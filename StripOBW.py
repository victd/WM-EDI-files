

import docx

from docx.shared import Pt

doc = docx.Document()

#style = doc.styles['Normal']
#font = style.font
#font.name = 'Courier New'
#font.size = Pt(10)


# CHANGE THIS TEXT FILE FOR THE DIVISION ----------------------------
# Refer to the TRUX EDI extract template for the fields in csv format
# for future revisions, see WM EDI extract
# third party company disclosure not needed as appears to be standard trux template
# trux disclosure
# conforms to a pattern yes no
# strips OBW but need to distinguish from Roll Off bins, keep the weights on RO bins
# the EDI extract format, each field, does not distinguish type of weight
# will need to join tables at the service level first
# we get feedback from Oakleaf, it will be 2-3 months behind as they process the data on their end
# some consolidated billing stores may have trouble distinguishing sites under the same storeID
# prefer to be under separate site billing
# active directory connect sync issue azure, kayak, stand up paddle
# divide and conquer to resolve cumulation of issues, folder mappings
# elevator and hvac service

f = open("WT.txt", "r")
f2 = open("WT-stripped.txt", "w")

Lines = f.readlines()

#currline = f.readline()
#doc.add_paragraph(currline)

#count = 0

for line in Lines:
    tempLine = line.split(',')
    lineType = tempLine[0]
    if "D" in lineType:
        Tdate = tempLine[1]
        Invoice = tempLine[2]
        Account = tempLine[3]
        Description = tempLine[4]
        Size = tempLine[5]
        Quantity = tempLine[6]
        Rate = tempLine[7]
        Tax = tempLine[8]
        OBweight = tempLine[9]
        Cost = tempLine[10]
        f2.write("D" + "," + Tdate + "," + Invoice + "," + Account + "," + Description + "," + Size + "," + Quantity + "," + Rate + "," + Tax + "," + "0.00" + "," + Cost + "," + "\n")
    else:
        f2.write(line)
	




f.close()
f2.close()


# CHANGE THIS DOCX FILE FOR THE DIVISION ----------------------------
# looks like only Toronto division has the mixed lines issue

#doc.save('WT-OBW.txt')




