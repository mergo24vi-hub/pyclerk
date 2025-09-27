import os

directory = 'scans'

def findscan(prefix,p):
    directory = prefix+'scans'
    for filename in os.listdir(directory):
        if filename.lower().startswith(p):
            return os.path.join(directory, filename)
    return None

def findava(prefix,p):
    directory = prefix+'avas'
    for filename in os.listdir(directory):
        if filename.lower().startswith(p):
            return os.path.join(directory, filename)
    return None

if __name__=='__main__':
    result = findscan('ua','бабенко')
    if result:
        print("Знайдено:", result)
    else:
        print("Файл не знайдено")