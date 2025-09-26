import os
directory = 'scans'

def finduascan(prefix):
    for filename in os.listdir(directory):
        if filename.lower().startswith(prefix):
            return os.path.join(directory, filename)
    return None

if __name__=='__main__':
    result = finduascan('бабенко')
    if result:
        print("Знайдено:", result)
    else:
        print("Файл не знайдено")