"""
Name: Sedem Quame Amekpewu
GitHub: @SedemQuame
Date: Saturday, Oct 24, 2020
Descriptiont: Script to shuffle people in the turntabl program into groups for presentation and group
              work.
"""

# import statements
import random
import xlsxwriter


# swapping elements positions in array
def swapArrElements(listOfNumbers, position1, position2):
    temp = listOfNumbers[position1]
    listOfNumbers[position1] = listOfNumbers[position2]
    listOfNumbers[position2] = temp
    return listOfNumbers

def printGroup(group):
    print(group)

# show groups
def showGroups(listOfNumbers, groupSize):
    j = 0
    group = []
    remainderMark = len(listOfNumbers) - (len(listOfNumbers) % groupSize)
    count = 0
    for number in listOfNumbers:
        group.append(number)
        j += 1
        if j == groupSize:
            # print group
            printGroup(group)
            group = []
            j = 0
        if count == remainderMark:
            print("@ the remainder mark")
            # group is equal to sub arr from index of remainder mark to array end
            group = listOfNumbers[remainderMark:]
            printGroup(group)
        count += 1

def createRowAndColHeaders(worksheet, numberOfGroups):
    row = 0
    col = 0
    colHeaders = [f'Group', f'Date', f'Topic', f'Members']
    # creating row headers
    for header in colHeaders:
        worksheet.write(row, col, header)
        col += 2
 
    #creating col numbers
    row = 2
    col = 0
    groupNumber = 'Group'
    i = 1
    x = numberOfGroups * 2
    while(row <= x):
        worksheet.write(row, col, f'Group #{i}')
        i += 1
        row += 2    

def convertArrToString(arr):
    nameList = ''
    for name in arr:
        nameList = nameList + name + ', '
    return nameList

def writeDataToWorkSheet(worksheet, topics, listOfNumbers, groupSize):
    j = 0
    row = 2
    col = 6
    group = []
    fieldMembers = ""
    remainderMark = len(listOfNumbers) - (len(listOfNumbers) % groupSize)
    count = 0
    i = 0
    for number in listOfNumbers:
        group.append(number)
        j += 1
        if j == groupSize:
            worksheet.write(row, col - 2, topics[i])
            i += 1
            worksheet.write(row, col, convertArrToString(group))
            row += 2
            group = []
            fieldMembers = ""
            j = 0
        if count == remainderMark:
            print("@ the remainder mark")
            # group is equal to sub arr from index of remainder mark to array end
            group = listOfNumbers[remainderMark:]
            worksheet.write(row, col - 2, topics[i])
            i += 1
            worksheet.write(row, col, convertArrToString(group))
            row += 2
        count += 1

# random input
listOfNumbers = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']
topics = ['algebra', 'math', 'english', 'rme', 'social studies', 'mas' 's','sd', '4', 'r']
print("Before shuffling list")
print(listOfNumbers)
print("\n")

# get size of group.
groupSize = int(input(f'Group size: '))

arrSize = len(listOfNumbers) - 1
# using fisher-yates algorithm.
lastUnShuffledPosition = arrSize
for number in listOfNumbers:
    randomIndex = random.randint(0, arrSize)
    listOfNumbers = swapArrElements(listOfNumbers, randomIndex, lastUnShuffledPosition)
    arrSize -= 1
    lastUnShuffledPosition -= 1

print("\nAfter shuffling list")
print(listOfNumbers)
print("\n")

showGroups(listOfNumbers, groupSize)

nameOfExcelFile = input(f'Name of excel file: ')
# writing to an excel file
workbook = xlsxwriter.Workbook(f'{nameOfExcelFile}.xlsx')

# add a new worksheet
worksheet = workbook.add_worksheet()

# write to a worksheet
createRowAndColHeaders(worksheet, 7)

#write group data to worksheet
writeDataToWorkSheet(worksheet, topics, listOfNumbers, groupSize)

workbook.close()