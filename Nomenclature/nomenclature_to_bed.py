#nomenclature to bed format
import re
import os

print("Covert nomenclature to bed")
end = ''
for line in iter(input, end):
    lo = re.search(r'(.*)\s(\d{1,2}|\w)(.*)[(](\d*)[_](\d*)[)](x|\s)(.*)', line)
    chromelo = lo.group(2)
    start=lo.group(4)
    end=lo.group(5)
    print("chr" + chromelo + "\t" + start + "\t" + end)
   
