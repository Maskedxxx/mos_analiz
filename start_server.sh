#!/bin/bash
# ============================================================================
# Полный запуск приложения: Docker-модели → API → UI → Tuna-туннель
# Использование: bash start_server.sh
# Остановка: Ctrl+C (все процессы завершатся автоматически)
# ============================================================================

set -e

PROJECT_DIR="/home/masked/projects/mos_analiz"
VENV="$PROJECT_DIR/.venv/bin/activate"
UI_DIR="$PROJECT_DIR/ui-demo"

# Логин/пароль для веб-интерфейса
export AUDIT_LOGIN="${AUDIT_LOGIN:-admin}"
export AUDIT_PASSWORD="${AUDIT_PASSWORD:-mos186124kva}"

# Docker-образ VLLM
VLLM_IMAGE="scitrera/dgx-spark-vllm:0.15.1-t4"
HF_CACHE="$HOME/.cache/huggingface"

# Порты
PORT_OCR=8010
PORT_LLM=8001
PORT_API=8080
PORT_UI=5173

# Массив PID фоновых процессов — для cleanup
PIDS=()

cleanup() {
    echo ""
    echo "═══════════════════════════════════════"
    echo "Остановка всех сервисов..."
    echo "═══════════════════════════════════════"
    # Убиваем фоновые процессы (tuna, uvicorn, vite)
    for pid in "${PIDS[@]}"; do
        kill "$pid" 2>/dev/null
    done
    # Останавливаем Docker-контейнеры
    docker rm -f paddleocr-vl-15 gpt-oss-120b 2>/dev/null
    echo "Все сервисы остановлены."
    exit 0
}
trap cleanup SIGINT SIGTERM

# ────────────────────────────────────────────────────────────────
# Функция ожидания готовности сервиса
# ────────────────────────────────────────────────────────────────
wait_for_port() {
    local port=$1
    local name=$2
    local timeout=${3:-300}
    local elapsed=0
    echo -n "  Ожидание $name (порт $port)..."
    while ! curl -s "http://localhost:$port" >/dev/null 2>&1; do
        sleep 3
        elapsed=$((elapsed + 3))
        if [ $elapsed -ge $timeout ]; then
            echo " ТАЙМАУТ (${timeout}с)!"
            return 1
        fi
        echo -n "."
    done
    echo " готов! (${elapsed}с)"
}

echo "═══════════════════════════════════════"
echo "  Запуск приложения МосМониторинг"
echo "═══════════════════════════════════════"
echo ""

# ────────────────────────────────────────────────────────────────
# Шаг 1: Docker-модели
# ────────────────────────────────────────────────────────────────
echo ">>> Шаг 1/4: Docker-модели"

# Проверяем, не запущены ли уже контейнеры
if docker ps --format '{{.Names}}' | grep -q "paddleocr-vl-15"; then
    echo "  PaddleOCR-VL-1.5 уже запущен"
else
    echo "  Запуск PaddleOCR-VL-1.5 (порт $PORT_OCR)..."
    docker run --rm -d --gpus all --network host --ipc=host \
      --name paddleocr-vl-15 \
      -v "$HF_CACHE:/root/.cache/huggingface" \
      "$VLLM_IMAGE" \
      vllm serve PaddlePaddle/PaddleOCR-VL-1.5 \
      --trust-remote-code \
      --gpu-memory-utilization 0.05 \
      --max-model-len 4096 \
      --max-num-batched-tokens 4096 \
      --no-enable-prefix-caching \
      --served-model-name PaddleOCR-VL-1.5 \
      --host 0.0.0.0 --port $PORT_OCR
fi

# Ждём PaddleOCR ПЕРЕД запуском LLM (иначе OOM при одновременной загрузке)
wait_for_port $PORT_OCR "PaddleOCR" 300

if docker ps --format '{{.Names}}' | grep -q "gpt-oss-120b"; then
    echo "  gpt-oss-120b уже запущен"
else
    echo "  Запуск gpt-oss-120b (порт $PORT_LLM)..."
    docker run --rm -d --gpus all --network host --ipc=host \
      --name gpt-oss-120b \
      -v "$HF_CACHE:/root/.cache/huggingface" \
      "$VLLM_IMAGE" \
      vllm serve openai/gpt-oss-120b \
      --trust-remote-code \
      --gpu-memory-utilization 0.60 \
      --max-model-len 96000 \
      --served-model-name openai/gpt-oss-120b \
      --host 0.0.0.0 --port $PORT_LLM
fi

# Ждём готовности LLM
wait_for_port $PORT_LLM "gpt-oss-120b" 600

echo ""

# ────────────────────────────────────────────────────────────────
# Шаг 2: API-сервер (uvicorn)
# ────────────────────────────────────────────────────────────────
echo ">>> Шаг 2/4: API-сервер"
source "$VENV"
cd "$PROJECT_DIR"

uvicorn api_server:app --host 0.0.0.0 --port $PORT_API &
PIDS+=($!)
wait_for_port $PORT_API "API-сервер" 60

echo ""

# ────────────────────────────────────────────────────────────────
# Шаг 3: UI (Vite dev server)
# ────────────────────────────────────────────────────────────────
echo ">>> Шаг 3/4: UI (Vite)"
cd "$UI_DIR"

npx vite --host 0.0.0.0 --port $PORT_UI &
PIDS+=($!)
sleep 3

echo ""

# ────────────────────────────────────────────────────────────────
# Шаг 4: Tuna-туннель (проброс UI наружу)
# ────────────────────────────────────────────────────────────────
echo ">>> Шаг 4/4: Tuna-туннель"

# Запускаем tuna и вытаскиваем публичную ссылку из вывода
TUNA_LOG=$(mktemp)
tuna http $PORT_UI --subdomain=mosaudit > "$TUNA_LOG" 2>&1 &
PIDS+=($!)

# Ждём появления ссылки в логе
TUNA_URL=""
for i in $(seq 1 15); do
    TUNA_URL=$(grep -oP 'https://\S+\.ru\.tuna\.am' "$TUNA_LOG" 2>/dev/null | head -1)
    if [ -n "$TUNA_URL" ]; then
        break
    fi
    sleep 1
done
rm -f "$TUNA_LOG"

echo ""
echo "═══════════════════════════════════════"
echo "  Все сервисы запущены!"
echo "═══════════════════════════════════════"
echo ""
echo "  Локально:   http://localhost:$PORT_UI"
if [ -n "$TUNA_URL" ]; then
echo "  Для коллег:  $TUNA_URL"
fi
echo ""
echo "  Логин:       $AUDIT_LOGIN"
echo "  Пароль:      $AUDIT_PASSWORD"
echo ""
echo "  Ctrl+C — остановить все сервисы"
echo "═══════════════════════════════════════"

# Ждём завершения (Ctrl+C → cleanup)
wait
