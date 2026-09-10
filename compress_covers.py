import os
from PIL import Image

# --- Настройки ---
COVERS_DIR = r"E:\BibCatalog\covers"
MAX_WIDTH = 800   # Максимальная ширина в пикселях
QUALITY = 75      # Качество JPEG (75-85 — хороший баланс) [citation:4]

def compress_images(directory):
    """Сжимает все JPG-изображения в папке, уменьшая их до MAX_WIDTH и перезаписывая оригиналы."""
    for root, _, files in os.walk(directory):
        for filename in files:
            if filename.lower().endswith(('.jpg', '.jpeg')):
                filepath = os.path.join(root, filename)
                try:
                    with Image.open(filepath) as img:
                        # Конвертируем в RGB, если необходимо
                        if img.mode in ('RGBA', 'P'):
                            img = img.convert('RGB')

                        # Изменяем размер, если изображение шире MAX_WIDTH
                        if img.width > MAX_WIDTH:
                            ratio = MAX_WIDTH / img.width
                            new_height = int(img.height * ratio)
                            img = img.resize((MAX_WIDTH, new_height), Image.LANCZOS)

                        # Сохраняем сжатое изображение, перезаписывая оригинал
                        img.save(filepath, 'JPEG', quality=QUALITY, optimize=True)
                        print(f"Сжато: {filename}")
                except Exception as e:
                    print(f"Ошибка при обработке {filepath}: {e}")

if __name__ == "__main__":
    compress_images(COVERS_DIR)
    print("\nГотово! Все обложки сжаты.")