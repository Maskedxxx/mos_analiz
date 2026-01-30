import os
import time
import shutil
import subprocess
from faster_whisper import WhisperModel

# ================== НАСТРОЙКИ ПУТЕЙ ==================
# ROOT_DIR — корневая папка, в которой лежат ПОДПАПКИ с видео.
# Пример: "/content/drive/MyDrive/Мои_вебинары"
ROOT_DIR = "/content/drive/MyDrive/видео_ВМЕСТЕ"  # <-- ЗДЕСЬ МЕНЯЕШЬ ПОД СЕБЯ

# ВАЖНО:
# Внутри ROOT_DIR у тебя лежат:
#   - подпапки с видео (любой структуры),
#   - мы ДОБАВИМ рядом служебные папки:
#       _transcripts  — сюда складываем результаты транскрибации
#       _processed    — сюда переносим успешно обработанные видео
#       _errors       — сюда переносим видео, на которых что-то сломалось

INPUT_DIR = ROOT_DIR                          # тут живут исходные папки с видео
OUTPUT_DIR = os.path.join(ROOT_DIR, "_transcripts")
PROCESSED_DIR = os.path.join(ROOT_DIR, "_processed")
ERROR_DIR = os.path.join(ROOT_DIR, "_errors")

# ================== НАСТРОЙКИ МОДЕЛИ ==================
MODEL_SIZE = "medium"                 # можешь поставить "medium" или "large-v3"
DEVICE = "cuda"                      # "cuda" или "cpu"
COMPUTE_TYPE = "float16" if DEVICE == "cuda" else "int8"

# Поддерживаемые форматы видео
VIDEO_EXTENSIONS = ('.mp4', '.mov', '.avi', '.mkv', '.ts', '.flv')


# ================== УТИЛИТЫ ==================
def log(msg: str):
    """Простой лог с таймстампом, чтобы понимать, что происходит."""
    ts = time.strftime('%Y-%m-%d %H:%M:%S')
    print(f"[{ts}] {msg}")


def setup_dirs():
    """Создаём служебные директории, если их ещё нет."""
    for d in [OUTPUT_DIR, PROCESSED_DIR, ERROR_DIR]:
        os.makedirs(d, exist_ok=True)
    log("Директории инициализированы.")


def extract_audio(video_path: str, audio_output_path: str) -> bool:
    """
    Извлекает аудио 16kHz mono WAV через FFmpeg.
    video_path        — полный путь до видео
    audio_output_path — полный путь до .wav
    """
    command = [
        "ffmpeg",
        "-i", video_path,
        "-vn",
        "-acodec", "pcm_s16le",
        "-ar", "16000",
        "-ac", "1",
        "-y",
        audio_output_path,
    ]
    try:
        subprocess.run(command, check=True, stderr=subprocess.PIPE)
        return True
    except subprocess.CalledProcessError as e:
        log(f"❌ Ошибка FFmpeg: {e.stderr.decode(errors='ignore')}")
        return False


def iter_video_files():
    """
    Рекурсивно обходит INPUT_DIR и возвращает:
    (full_path, rel_path, filename)

    full_path — полный путь до файла
    rel_path  — путь относительно INPUT_DIR (например: "folder1/video.mp4")
    filename  — только имя файла ("video.mp4")
    """
    # Чтобы не заходить в служебные папки (_transcripts, _processed, _errors),
    # мы отфильтруем их на верхнем уровне.
    skip_dirs = {
        os.path.basename(OUTPUT_DIR),
        os.path.basename(PROCESSED_DIR),
        os.path.basename(ERROR_DIR),
    }

    for root, dirs, files in os.walk(INPUT_DIR):
        # На верхнем уровне выкидываем служебные директории из обхода
        if root == INPUT_DIR:
            dirs[:] = [d for d in dirs if d not in skip_dirs]

        for f in files:
            if not f.lower().endswith(VIDEO_EXTENSIONS):
                continue

            full_path = os.path.join(root, f)
            # относительный путь от корневой input-папки
            rel_path = os.path.relpath(full_path, INPUT_DIR)
            yield full_path, rel_path, f


def process_file(full_path: str, rel_path: str, filename: str, model: WhisperModel):
    """
    Обработка одного видео:
    - извлекаем аудио
    - транскрибируем
    - сохраняем текст
    - переносим видео в processed/errors

    full_path — полный путь до исходного видео
    rel_path  — относительный путь от INPUT_DIR (например: "folder1/video.mp4")
    filename  — имя файла ("video.mp4")
    """
    # относительная директория (подпапка) от INPUT_DIR, если есть
    relative_dir = os.path.dirname(rel_path)  # "" или "folder1/subfolder2"
    base_name = os.path.splitext(filename)[0]

    # Папка для результата конкретного видео:
    #   _transcripts/[relative_dir]/[video_name]/
    target_folder = os.path.join(OUTPUT_DIR, relative_dir, base_name)
    os.makedirs(target_folder, exist_ok=True)

    audio_path = os.path.join(target_folder, f"{base_name}.wav")
    txt_path = os.path.join(target_folder, f"{base_name}.txt")

    log(f"🎬 Обработка файла: {full_path}")
    log(f"📁 Папка результата: {target_folder}")

    # 1. Извлекаем аудио
    if not extract_audio(full_path, audio_path):
        # Если FFmpeg упал — переносим видео в errors, сохраняя структуру
        error_video_dir = os.path.join(ERROR_DIR, relative_dir)
        os.makedirs(error_video_dir, exist_ok=True)
        shutil.move(full_path, os.path.join(error_video_dir, filename))
        log(f"📦 Видео перемещено в errors: {os.path.join(error_video_dir, filename)}")
        return

    # 2. Транскрибируем
    log("🎙 Начало транскрибации...")
    start_time = time.time()

    try:
        segments, info = model.transcribe(
            audio_path,
            beam_size=5,
            language="ru",
            vad_filter=True,          # вырезать тишину/шум
            vad_parameters={"min_silence_duration_ms": 500},
            condition_on_previous_text=True,
        )

        full_text_lines = []
        for segment in segments:
            # Формат: [MM:SS - MM:SS] Текст
            start_fmt = time.strftime('%M:%S', time.gmtime(segment.start))
            end_fmt = time.strftime('%M:%S', time.gmtime(segment.end))
            line = f"[{start_fmt} - {end_fmt}] {segment.text}"
            full_text_lines.append(line)

        # 3. Сохраняем текст
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write("\n".join(full_text_lines))

        elapsed = time.time() - start_time
        log(f"✅ Транскрибация завершена за {elapsed:.2f} сек. Результат: {txt_path}")

        # 4. Перемещаем исходное видео в processed, сохраняя структуру
        processed_video_dir = os.path.join(PROCESSED_DIR, relative_dir)
        os.makedirs(processed_video_dir, exist_ok=True)
        shutil.move(full_path, os.path.join(processed_video_dir, filename))
        log(f"📦 Видео перемещено в processed: {os.path.join(processed_video_dir, filename)}")

    except Exception as e:
        log(f"❌ Ошибка Whisper: {e}")
        # Переносим исходник в errors
        error_video_dir = os.path.join(ERROR_DIR, relative_dir)
        os.makedirs(error_video_dir, exist_ok=True)
        shutil.move(full_path, os.path.join(error_video_dir, filename))
        log(f"📦 Видео перемещено в errors: {os.path.join(error_video_dir, filename)}")


def main_loop():
    """Основной цикл: периодически сканирует папку и обрабатывает новые видео."""
    setup_dirs()
    log(f"🚀 Сервис запущен.")
    log(f"📂 Корневая папка с видео: {INPUT_DIR}")
    log(f"🧠 Модель: {MODEL_SIZE} (device={DEVICE}, compute={COMPUTE_TYPE})")

    # Загружаем модель один раз
    try:
        model = WhisperModel(MODEL_SIZE, device=DEVICE, compute_type=COMPUTE_TYPE)
    except Exception as e:
        log(f"❌ Ошибка загрузки модели (GPU?): {e}")
        log("⚠️ Переключаюсь на CPU...")
        model = WhisperModel(MODEL_SIZE, device="cpu", compute_type="int8")

    # Бесконечный цикл — подходит под "демон".
    # В Colab можно просто запустить ячейку и оставить.
    while True:
        any_files = False

        for full_path, rel_path, filename in iter_video_files():
            any_files = True
            process_file(full_path, rel_path, filename, model)

        if not any_files:
            log("🔎 Новых видео не найдено. Ждём...")

        # Пауза перед следующим проходом
        time.sleep(10)


if __name__ == "__main__":
    main_loop()
