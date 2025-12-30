"""
Скрипт для создания иконки принтера для приложения Kyocera Scan Selector
"""

try:
    from PIL import Image, ImageDraw
    import os

    # Создаем изображение 64x64
    size = 64
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Цвета
    printer_color = (45, 95, 141)  # Синий цвет
    paper_color = (255, 255, 255)  # Белый
    shadow_color = (100, 100, 100, 100)  # Серая тень

    # Рисуем корпус принтера
    draw.rectangle([10, 25, 54, 50], fill=printer_color, outline=(30, 70, 110), width=2)

    # Рисуем бумагу
    draw.rectangle([15, 15, 49, 30], fill=paper_color, outline=(200, 200, 200), width=1)

    # Рисуем линии на бумаге
    for y in [19, 22, 25]:
        draw.line([19, y, 45, y], fill=(150, 150, 150), width=1)

    # Рисуем лоток
    draw.rectangle([12, 48, 52, 54], fill=(35, 85, 131), outline=(25, 65, 111), width=1)

    # Рисуем кнопки
    draw.ellipse([42, 32, 48, 38], fill=(0, 200, 100))
    draw.ellipse([42, 40, 48, 46], fill=(200, 50, 50))

    # Сохраняем в разных размерах для .ico
    icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    images = []

    for icon_size in icon_sizes:
        resized = img.resize(icon_size, Image.Resampling.LANCZOS)
        images.append(resized)

    # Сохраняем как .ico
    script_dir = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(script_dir, 'printer.ico')

    images[0].save(icon_path, format='ICO', sizes=[(s.width, s.height) for s in images], append_images=images[1:])

    print(f"✓ Иконка успешно создана: {icon_path}")

except ImportError:
    print("Ошибка: Требуется библиотека Pillow")
    print("Установите: pip install Pillow")
except Exception as e:
    print(f"Ошибка при создании иконки: {e}")
