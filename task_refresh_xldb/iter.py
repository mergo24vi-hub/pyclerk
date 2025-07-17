#G:\.shortcut-targets-by-id\1Pcnp8gnqT8NS3Zl5AOanpcBmZLHuuv5I\РО робоча\ОСОБИСТІ ПАПКИ
import os
root_dir = r'G:\.shortcut-targets-by-id\1Pcnp8gnqT8NS3Zl5AOanpcBmZLHuuv5I\РО робоча\ОСОБИСТІ ПАПКИ\\'
i=0
for folder in os.listdir(root_dir):
    folder_path = os.path.join(root_dir, folder)

    if os.path.isdir(folder_path):
        print(folder_path)
        i=i+1
        if i>399:
           break
        matching_files = []

        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if not os.path.isfile(file_path):
                continue
            #print(file_path)
            # Перевіряємо, що це файл .doc з 'сдд' у назві (незалежно від регістру)
            if file.lower().endswith('.docx') and 'сдд' in file.lower():
                matching_files.append(file)

        if matching_files:
            total = (len(matching_files))
            # Сортуємо за часом зміни
            matching_files.sort(key=lambda f: os.path.getmtime(os.path.join(folder_path, f)))
            print(f"У папці '{folder}' всього сдд {total} вибрано: {matching_files[-1]}")
            print('--------')

