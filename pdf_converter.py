import os
import time
import psutil
from pathlib import Path
from docx2pdf import convert

def cnv (conv_path: str):
    def kill_win_process():
        cleaned = False
    for proc in psutil.process_iter(['name']):
        try:
            # Проверяем имя процесса (в Windows это WINWORD.EXE)
            if proc.info['name'] and proc.info['name'].upper() == 'WINWORD.EXE':
                proc.kill()  # Жестко убиваем процесс
                cleaned = True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

    root_dir = Path(conv_path).resolve()
    #output_dir = input_dir / 'Output'
    converted_dirs = []
    for dirpath, dirnames, filenames in os.walk(root_dir):
        if 'Output' in Path(dirpath).parts:
            continue

        current_dir = Path(dirpath)
        output_dir = current_dir / 'Output'
        output_dir.mkdir(parents=True, exist_ok=True)
        #convert(str(current_dir), str(output_dir))
        print(f'Конвертирую директорию: {current_dir}')

        try:
            convert(str(current_dir), str(output_dir))
            converted_dirs.append(str(output_dir))
            time.sleep(2)
            #kill_win_process
        except Exception as e:
            print(f'Возникла ошибка  в {output_dir} : {e}')
        
        

    print('\nКонвертирование выполнено. Файлы лежат в следующих директориях: ')
    for i in converted_dirs:
        print(i)
    
conv_str = input('Введите путь к директории: ')
cnv(conv_str)