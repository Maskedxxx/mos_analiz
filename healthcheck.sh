#!/bin/bash
# Скрипт проверки здоровья сервисов MosAudit
# Запускается по cron каждые 5 минут
# Логирует в logs_result/healthcheck.log

TIMESTAMP=$(date "+%Y-%m-%d %H:%M:%S")
SUDO_PASS="Lqv#*El9"
FAILURES=()

# --- Проверка API ---
API_RESPONSE=$(curl -s --max-time 10 http://localhost:8080/api/health 2>/dev/null)
API_STATUS=$(echo "$API_RESPONSE" | python3 -c "import sys,json; print(json.load(sys.stdin).get(\"status\",\"fail\"))" 2>/dev/null)

if [ "$API_STATUS" != "ok" ]; then
    FAILURES+=("api")
    echo "[$TIMESTAMP] FAIL api — restart mosaudit-api"
    echo "$SUDO_PASS" | sudo -S systemctl restart mosaudit-api 2>/dev/null
    sleep 3
    # Повторная проверка
    RECHECK=$(curl -s --max-time 10 http://localhost:8080/api/health 2>/dev/null | python3 -c "import sys,json; print(json.load(sys.stdin).get(\"status\",\"fail\"))" 2>/dev/null)
    if [ "$RECHECK" = "ok" ]; then
        echo "[$TIMESTAMP] OK — mosaudit-api восстановлен"
    else
        echo "[$TIMESTAMP] CRITICAL — mosaudit-api НЕ восстановлен"
    fi
fi

# --- Проверка Nginx/UI ---
UI_CODE=$(curl -s -o /dev/null -w "%{http_code}" --max-time 5 http://localhost:5173 2>/dev/null)

if [ "$UI_CODE" != "200" ]; then
    FAILURES+=("nginx")
    echo "[$TIMESTAMP] FAIL nginx — HTTP $UI_CODE — restart nginx"
    echo "$SUDO_PASS" | sudo -S systemctl restart nginx 2>/dev/null
    sleep 2
    RECHECK_UI=$(curl -s -o /dev/null -w "%{http_code}" --max-time 5 http://localhost:5173 2>/dev/null)
    if [ "$RECHECK_UI" = "200" ]; then
        echo "[$TIMESTAMP] OK — nginx восстановлен"
    else
        echo "[$TIMESTAMP] CRITICAL — nginx НЕ восстановлен"
    fi
fi

# --- Проверка Tuna ---
TUNA_CODE=$(curl -s -o /dev/null -w "%{http_code}" --max-time 10 https://mosaudit.ru.tuna.am 2>/dev/null)

if [ "$TUNA_CODE" != "200" ]; then
    FAILURES+=("tuna")
    echo "[$TIMESTAMP] FAIL tuna — HTTP $TUNA_CODE — restart mosaudit-tuna"
    echo "$SUDO_PASS" | sudo -S systemctl restart mosaudit-tuna 2>/dev/null
    sleep 5
    RECHECK_TUNA=$(curl -s -o /dev/null -w "%{http_code}" --max-time 10 https://mosaudit.ru.tuna.am 2>/dev/null)
    if [ "$RECHECK_TUNA" = "200" ]; then
        echo "[$TIMESTAMP] OK — tuna восстановлен"
    else
        echo "[$TIMESTAMP] CRITICAL — tuna НЕ восстановлен"
    fi
fi

# --- Проверка моделей на Spark (из ответа API health) ---
if [ "$API_STATUS" = "ok" ]; then
    PADDLE_STATUS=$(echo "$API_RESPONSE" | python3 -c "import sys,json; print(json.load(sys.stdin)[\"services\"][\"paddleocr\"][\"status\"])" 2>/dev/null)
    LLM_STATUS=$(echo "$API_RESPONSE" | python3 -c "import sys,json; print(json.load(sys.stdin)[\"services\"][\"llm\"][\"status\"])" 2>/dev/null)

    if [ "$PADDLE_STATUS" != "ok" ]; then
        FAILURES+=("paddleocr")
        echo "[$TIMESTAMP] WARN paddleocr на Spark недоступен (status: $PADDLE_STATUS)"
    fi
    if [ "$LLM_STATUS" != "ok" ]; then
        FAILURES+=("llm")
        echo "[$TIMESTAMP] WARN llm на Spark недоступен (status: $LLM_STATUS)"
    fi
fi

# --- Итог ---
if [ ${#FAILURES[@]} -eq 0 ]; then
    echo "[$TIMESTAMP] OK — все сервисы работают"
fi
