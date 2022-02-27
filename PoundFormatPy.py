import openpyxl

path = "PoundFormat.xlsx"
file = openpyxl.load_workbook(path)
sheet = file.active

weights = []
stones = []
pounds = []

for i in range(2, 54):

    cell = sheet.cell(row = i, column = 5)

    if cell.value != None:
        weights.append(cell.value)

for weight in weights:
    weight_split = weight.split()

    stone = weight_split[0]

    if len(weight_split) != 1:
        pound = weight_split[1]
        pounds.append(pound)
    else:
        pounds.append('0lb')

    stones.append(stone)
    

print(f'Weights : {weights}')
print("\n")
print(f'Stones : {stones}')
print("\n")
print(f'Pounds : {pounds}')

num_stones = []
num_pounds = []

for pound in pounds:
    numeric_filter = filter(str.isdigit, pound)
    pound = "".join(numeric_filter)
    pound = int(pound)
    num_pounds.append(pound)

for stone in stones:    #stone = '00st'
    numeric_filter = filter(str.isdigit, stone)
    stone = "".join(numeric_filter)
    stone = int(stone) * 14
    num_stones.append(stone)

print("\n")
print(num_pounds)
print(num_stones)

total = []

for i in range(0, len(num_stones)):
    total.append(num_pounds[i] + num_stones[i])

print("\n")
print(total)

total_index = 0

for i in range(len(total)):
    sheet[f"F{i+2}"] = f"{total[total_index]}lb"
    total_index += 1

file.save(filename="PoundFormat.xlsx")
