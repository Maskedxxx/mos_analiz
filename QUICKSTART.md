# Быстрый запуск

4 терминала, запускать **строго по порядку**. Каждый следующий шаг — только после "Uvicorn running" в предыдущем.

> **ВАЖНО:** OCR и LLM контейнеры нельзя запускать одновременно! На UMA-памяти DGX Spark vLLM профилирует всю GPU-память при старте. Если оба стартуют параллельно — один из них упадёт с ошибкой "No available memory for the cache blocks".

## 1. OCR-модель (терминал 1)

```bash
docker run --rm --gpus all --network host --ipc=host \
  --name paddleocr-vl-15 \
  -v ~/.cache/huggingface:/root/.cache/huggingface \
  scitrera/dgx-spark-vllm:0.15.1-t4 \
  vllm serve PaddlePaddle/PaddleOCR-VL-1.5 \
  --trust-remote-code \
  --gpu-memory-utilization 0.05 \
  --max-model-len 4096 \
  --max-num-batched-tokens 4096 \
  --no-enable-prefix-caching \
  --served-model-name PaddleOCR-VL-1.5 \
  --host 0.0.0.0 --port 8000
```

## 2. LLM-модель (терминал 2)

```bash
docker run --rm --gpus all --network host --ipc=host \
  --name gpt-oss-120b \
  -v ~/.cache/huggingface:/root/.cache/huggingface \
  scitrera/dgx-spark-vllm:0.15.1-t4 \
  vllm serve openai/gpt-oss-120b \
  --trust-remote-code \
  --gpu-memory-utilization 0.70 \
  --max-model-len 96000 \
  --served-model-name openai/gpt-oss-120b \
  --host 0.0.0.0 --port 8001
```

## 3. API-сервер (терминал 3)

```bash
cd ~/projects/mos_analiz
bash start_server.sh
# → http://localhost:8080
# Логин: admin / admin
```

## 4. UI (терминал 4)

```bash
cd ~/projects/mos_analiz/ui-demo
npm run dev
# → http://localhost:5173
```

## Проверка готовности

```bash
curl http://localhost:8000/v1/models   # OCR
curl http://localhost:8001/v1/models   # LLM
curl http://localhost:8080/api/types   # API
```

## Остановка

- Ctrl+C в каждом терминале
- Или: `docker rm -f paddleocr-vl-15 gpt-oss-120b`

## CLI-аудит (без UI)

```bash
source .venv/bin/activate
python run_audit.py --doc-type prikaz_ppu --target test_docs/prikaz_ppu/файл.docx
python run_audit.py --list-types
```
