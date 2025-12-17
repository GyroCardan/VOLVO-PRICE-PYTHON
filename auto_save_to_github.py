"""
Автоматическое сохранение изменений в git и отправка на GitHub.
Запускай этот скрипт периодически или перед завершением работы.
"""
import subprocess
import sys
import os
from datetime import datetime

def run_command(cmd, check=True):
    """Выполняет команду и возвращает результат."""
    try:
        result = subprocess.run(
            cmd, 
            shell=True, 
            capture_output=True, 
            text=True, 
            encoding='utf-8',
            check=check
        )
        return result.stdout.strip(), result.stderr.strip(), result.returncode
    except subprocess.CalledProcessError as e:
        return e.stdout.strip(), e.stderr.strip(), e.returncode

def auto_save():
    """Автоматически сохраняет все изменения в git и отправляет на GitHub."""
    print(f"🔄 Автосохранение: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)
    
    # Проверяем статус
    stdout, stderr, code = run_command("git status --short", check=False)
    
    if not stdout and code == 0:
        print("✅ Нет изменений для сохранения")
        return
    
    # Добавляем все изменения
    print("📦 Добавляю изменения...")
    stdout, stderr, code = run_command("git add volvo_price_txt_processing.py .gitignore", check=False)
    if code != 0:
        print(f"⚠️ Ошибка при добавлении: {stderr}")
        return
    
    # Проверяем, есть ли что коммитить
    stdout, stderr, code = run_command("git diff --cached --quiet", check=False)
    if code == 0:
        print("ℹ️ Нет изменений для коммита")
        return
    
    # Создаем коммит
    commit_message = f"Auto-save: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    print(f"💾 Создаю коммит: {commit_message}")
    stdout, stderr, code = run_command(f'git commit -m "{commit_message}"', check=False)
    if code != 0:
        print(f"⚠️ Ошибка при коммите: {stderr}")
        return
    
    print("✅ Коммит создан")
    
    # Отправляем на GitHub
    print("🚀 Отправляю на GitHub...")
    stdout, stderr, code = run_command("git push origin main", check=False)
    if code == 0:
        print("✅ Успешно отправлено на GitHub!")
        print(f"   Репозиторий: https://github.com/GyroCardan/VOLVO-PRICE-PYTHON")
    else:
        print(f"⚠️ Ошибка при отправке: {stderr}")
        print("   Попробуй выполнить вручную: git push origin main")
    
    print("=" * 60)

if __name__ == "__main__":
    try:
        auto_save()
    except KeyboardInterrupt:
        print("\n⛔ Прервано пользователем")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Ошибка: {e}")
        sys.exit(1)


