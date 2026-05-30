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

    # 1. Создаём venv
    if not os.path.exists(venv_path):
        print("🔄 Создание виртуального окружения...")
        subprocess.check_call([sys.executable, "-m", "venv", ".venv"])
        print("✅ Виртуальное окружение создано")
    else:
        print("✅ Виртуальное окружение уже существует")

    if not os.path.exists(venv_python):
        print("❌ Ошибка: python в .venv не найден!")
        sys.exit(1)

    # 2. Обязательные пакеты (критичны для запуска)
    essential = ["customtkinter", "pandas", "fuzzywuzzy", "python-Levenshtein", "openpyxl"]
    print(f"📦 Установка обязательных библиотек: {', '.join(essential)}...")
    try:
        subprocess.check_call([venv_python, "-m", "pip", "install"] + essential)
    except Exception as e:
        print(f"❌ Ошибка установки основных пакетов: {e}")
        sys.exit(1)

    # 3. Опциональный пакет (часто падает на Windows из-за отсутствия MSVC)
    print("📦 Попытка установки python-Levenshtein (ускоряет fuzzywuzzy)...")
    result = subprocess.run([venv_python, "-m", "pip", "install", "python-Levenshtein"],
                            capture_output=True, text=True)
    if result.returncode == 0:
        print("✅ python-Levenshtein установлен успешно")
    else:
        print("⚠️ python-Levenshtein не установлен (требуется C++ компилятор).")
        print("   fuzzywuzzy будет работать в стандартном режиме (предупреждение можно игнорировать).")

    print("\n✅ Все необходимые зависимости установлены!")

if __name__ == "__main__":
    main()