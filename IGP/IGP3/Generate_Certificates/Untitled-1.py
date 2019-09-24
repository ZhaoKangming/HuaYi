doclist = []
with open('D:\\doclist.txt', 'r') as txt:
    for line in txt:
        doclist.append(list(line.strip('\n').split(',')))

print(doclist)