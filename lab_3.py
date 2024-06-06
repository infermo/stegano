symbols = ['!', ';', ':', ',', '.', '?']

def main(file_name):
    file_text = open(file_name, "r")
    lines = file_text.readlines()

    bit_arrays = []
    index_array = 0

    for line in lines:
        bit_arrays.append([])
        for i in range(len(line)):
            if line[i] in symbols:
                bit_arrays[index_array].append(1)
            else:
                bit_arrays[index_array].append(0)
        #print(bit_arrays[index_array])
        index_array += 1
    for i in range(len(bit_arrays)):
        if len(bit_arrays[i]) % 2 == 0:
            bit_arrays[i] = [0]
    #print(bit_arrays)

    lines_yes = []
    lines_no = []
    index_str = 0
    for bits in bit_arrays:
        result = 0
        for bit in bits:
            result |= bit
        if result == 1:
            print(f"Строка {index_str+1}: ДА - {lines[index_str][:-1]}")
            lines_yes.append(lines[index_str])
        else:
            print(f"Строка {index_str+1}: НЕТ - {lines[index_str][:-1]}")
            lines_no.append(lines[index_str])
        index_str += 1

    print(f'\nДа: {len(lines_yes)}\nНет: {len(lines_no)}\n')

if __name__ == '__main__':
    print("Метод скрытия с помощью пунктуации:")
    main("laba3_2.txt")
