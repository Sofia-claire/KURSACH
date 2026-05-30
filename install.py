import os
import sys
import subprocess
import platform

def get_venv_python():
    if platform.system() == "Windows":
        return os.path.join(".venv", "Scripts", "python.exe")
    return os.path.join(".venv", "bin", "python")

def main():
    venv_path = os.path.join(os.getcwd(), ".venv")
    venv_python = get_venv_python()

    # 1. Создаём venv, если нет
    if not os.path.exists(venv_path):
        print("🔄 Создание виртуального окружения...")
        subprocess.check_call([sys.executable, "-m", "venv", ".venv"])
        print("✅ Виртуальное окружение создано")
    else:
        print("✅ Виртуальное окружение уже существует")

    if not os.path.exists(venv_python):
        print("❌ Ошибка: python в .venv не найден!")
        sys.exit(1)

    # 2. Обязательные пакеты
    # ✅ ДОБАВЛЕНО: Pillow (для иконок/PIL) и pystray (для трея)
    essential = [
        "customtkinter", 
        "pandas", 
        "fuzzywuzzy", 
        "openpyxl", 
        "Pillow", 
        "pystray"
    ]
    
    print(f"📦 Установка библиотек: {', '.join(essential)}...")
    try:
        subprocess.check_call([venv_python, "-m", "pip", "install"] + essential)
    except Exception as e:
        print(f"❌ Ошибка установки пакетов: {e}")
        sys.exit(1)

    # 3. Опционально: ускорение для fuzzywuzzy
    print("📦 Попытка установки python-Levenshtein (ускорение)...")
    result = subprocess.run([venv_python, "-m", "pip", "install", "python-Levenshtein"],
                            capture_output=True, text=True)
    if result.returncode == 0:
        print("✅ python-Levenshtein установлен")
    else:
        print("⚠️ python-Levenshtein не установлен (это нормально, fuzzywuzzy будет работать чуть медленнее).")

    print("\n✅ Все необходимые зависимости установлены!")

if __name__ == "__main__":
    main()