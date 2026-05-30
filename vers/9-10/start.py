import os
import sys
import subprocess
import platform

def get_venv_python():
    if platform.system() == "Windows":
        return os.path.join(".venv", "Scripts", "python.exe")
    return os.path.join(".venv", "bin", "python")

def check_dependencies():
    """Проверяет наличие критичных библиотек внутри venv"""
    venv_python = get_venv_python()
    if not os.path.exists(venv_python):
        return False
    
    try:
        subprocess.check_call(
            [venv_python, "-c", "import customtkinter, pandas, fuzzywuzzy, Levenshtein, openpyxl"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        return True
    except subprocess.CalledProcessError:
        return False

def main():
    if not os.path.exists("main.py"):
        print("❌ main.py не найден в текущей директории!")
        sys.exit(1)

    print("=" * 60)
    print("🔧 СИСТЕМА ЗАПУСКА ПРОЕКТА")
    print("=" * 60)

    venv_python = get_venv_python()

    # Проверка и однократный запуск установки
    if not os.path.exists(venv_python) or not check_dependencies():
        print("⚠️ Зависимости отсутствуют. Запуск установки...")
        install_result = subprocess.run([sys.executable, "install.py"])
        if install_result.returncode != 0:
            print("❌ Установка завершилась с ошибкой. Проверьте вывод выше.")
            sys.exit(1)
    else:
        print("✅ Окружение и зависимости готовы.")

    # Финальная проверка перед запуском main.py
    if not check_dependencies():
        print("❌ Зависимости так и не удалось установить. Проверьте логи install.py")
        sys.exit(1)

    print("\n🚀 Запуск main.py...")
    print("=" * 60)
    subprocess.run([venv_python, "main.py"])

if __name__ == "__main__":
    main()