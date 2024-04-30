
def hello(text):
    print(f'Функция "{text}"')

print('Привет Мир')

if __name__ == '__main__':
    my_file = open("BabyFile.txt", "w+")
    my_file.write("Привет, файл!")
    my_file.close()