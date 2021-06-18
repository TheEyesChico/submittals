import re
from pandas.core.base import DataError
from text_extraction import *
import pandas as pd
import pickle
import openpyxl
import os
import csv

def read_data(txtfile):

    with open(txtfile, "r",encoding="utf-8") as f:
        list_data=f.readlines()
        str_data=" ".join(list_data)

    f.close()
    return list_data,str_data

def PartsDivision(str_data):

    parts = re.findall("PART[\s]{1,3}[1-3]", str_data)

    part1_start = re.search(parts[0], str_data).start()
    part2_start = re.search(parts[1], str_data).start()
    part3_start = re.search(parts[2], str_data).start()

    part1=str_data[part1_start:part2_start]
    part2=str_data[part2_start:part3_start]
    part3=str_data[part3_start:]

    return [part1,part2,part3]

def parts_list(list_data):
    l1=[]
    for i,j in enumerate(list_data):
        if re.search("PART[\s]{1,3}[123]", j):
            l1.append(i)

    part1_list=list_data[l1[0]:l1[1]]
    part2_list=list_data[l1[1]:l1[2]]
    part3_list=list_data[l1[2]:]

    return [part1_list,part2_list,part3_list]

def SectionHeading(parts):

    part1,part2,part3=parts

    section_heading=r"\b[123][.]\d{1,2}[\s]*[A-Za-z-\t ():),/.']+\b"
    
    part1_section_ = re.findall(section_heading, part1)
    part2_section_ = re.findall(section_heading, part2, re.M)
    part3_section_ = re.findall(section_heading, part3, re.M)

    part1_section=[i for i in part1_section_ if i.isupper()]
    part2_section=[i for i in part2_section_ if i.isupper()]
    part3_section=[i for i in part3_section_ if i.isupper()]

    # print(part1_section)
    return [part1_section,part2_section,part3_section]


def map_sections(multiple_parts):

    mapping={}

    for part in multiple_parts:
        for section in part:
            divide=section.split(" ")
            section_num=divide[0]
            section_name=" ".join(divide[1:]).strip()
            mapping[section_num]=section_name

    # print(mapping)
    return mapping

def sub_section(three_parts,headings):
    '''

    FOR REFERENCE
    re.sub(r"\s*\n\s*", r"\n", "\n".join(indiviual_list))
    r"(?m)(?P<Section>[A-Z]+\. .*(?:\n(?!\d|[A-Z]+\.).*)*)"
    r"\b([A-Z]\s*\.\s*)\b"

    '''

    part1_list,part2_list,part3_list=three_parts
    part1_heading,part2_heading,part3_heading=headings


    for sub_list,sub_heading in zip([part1_list],[part1_heading]):
        l1=[]
        # print(sub_list)
        # print("="*145)

        # print(sub_heading)
        # print(sub_list)
        # print("="*145)

        for new_search in sub_heading:

            for i,ok in enumerate(sub_list):
                # print(new_search)
                if new_search in ok.lstrip():
                    # print("String ->",ok.lstrip())
                    # print("Search ->",new_search)
                    l1.append(i)
                    # print(l1)
        l1.append(len(sub_list))
        
        master={}

        for i in range(len(l1)):

            try:
                
                indiviual_list = sub_list[l1[i]:l1[i + 1]]

                # print(indiviual_list)

                # print(indiviual_list[0])
                section_found = False
                in_section = False
                result = {}
                items = []
                current_section = ''
                for s in indiviual_list:
                    if not s.strip():
                        continue
                    if re.match(r'[A-Z]\s*\.\s*\b', s):
                        in_section = True
                        if current_section: # not the first time found
                            result[current_section] = items
                            items = []
                        current_section = s.rstrip()
                        section_found = True
                    elif section_found and re.match(r'\d{1,2}\s*\.', s):
                        in_section = False
                        items.append(s.rstrip())
                    elif section_found:
                        if in_section:
                            current_section += f' {s.rstrip()}'
                        else:
                            if items:
                                items[-1] += f' {s.rstrip()}'

                if current_section:
                    result[current_section] = items
                    master[indiviual_list[0].rstrip()]=result
                # print(result)
                # print("="*140)
            except IndexError as e:
                pass
        
        # print(master)
        # print("\n")
        # sec_find="1.1  SUMMARY"
        # print(master[sec_find])
    # print(master)
    return master

class Entries:

    def __init__(self, list_data,str_data):
        self.list_data=list_data
        self.str_data=str_data
        
    def spec_number(self):
        regex=r'SECTION\s*\d{2}\s*\d{2}\s*\d{2}'
        spec_num=re.search("\d[0-9\s]+",re.search(regex,self.str_data,re.IGNORECASE).group()).group()
        # print(section_name[:2])
        return spec_num.strip()

    def division(self,spec_number):

        master=pd.read_excel('master.xlsx',sheet_name="Sheet1")
        if spec_number[0]=='0':
            division=int(spec_number[1])
        else:
            division=int(spec_number[:2])

        division=spec_number[:2]+" - "+master.loc[master['Specification Group'] == division, 'Name'].iloc[0]
        return division

    def spec_name(self):

        regex=r'SECTION\s*\d{2}\s*\d{2}\s*\d{2}[i]?'

        _,start=re.search(regex,self.str_data,re.IGNORECASE).span()
        end,_=re.search("PART",self.str_data).span()
        spec_name=self.str_data[start:end].strip()

        return spec_name

    def test_reports(self,data):
        
        for key,value in data.items():
            
            if re.search('SUBMITTAL',key,re.I) :
                print(key)

        return 

    def search(self,word,data):
        
        sub_section={}
        submittals=True
        for key,value in data.items():
            
            if re.search('WARRANTY',key,re.I) or re.search('WARRANTIES',key,re.I):
                submittals=False
                found=False
                
                for i,j in value.items():
                    
                    manufacturer=re.search("[A-Z]\s*[.]\s*Manufacture([rs’]?)+\s*Warrant(y|ies)",i,re.I)
                    installer=re.search("[A-Z]\s*[.]\s*Installe([rs’]?)+\s*Warrant(y|ies)",i,re.I)
                    special=re.search("[A-Z]\s*[.]\s*Special([s’]?)+\s*Warrant(y|ies)",i,re.I)

                    if manufacturer or installer or special:

                        found=True
                        
                        if (j == []) :
                            sub_sec=key.split()[0]+"/"+i.split(" ", 1)[0].replace(".", "")
                            sub_section[sub_sec]=i.split(" ", 1)[1].strip()

                        else:
                            for item in j:
                                sub_sec=key.split()[0]+"/"+i.split(" ", 1)[0].replace(".", "")
                                sub_sec=sub_sec+"/"+item.split(" ", 1)[0].replace(".", "")
                                sub_section[sub_sec]=item.split(" ", 1)[1].strip()

                if not found:
                    
                    for i,j in value.items():

                        if (j == []) :
                            sub_sec=key.split()[0]+"/"+i.split(" ", 1)[0].replace(".", "")
                            sub_section[sub_sec]=i.split(" ", 1)[1].strip()

                        else:
                            for item in j:
                                sub_sec=key.split()[0]+"/"+i.split(" ", 1)[0].replace(".", "")
                                sub_sec=sub_sec+"/"+item.split(" ", 1)[0].replace(".", "")
                                sub_section[sub_sec]=item.split(" ", 1)[1].strip()
                                # print(sub_sec)
                                # print(sub_section)
        
        if submittals:
            
            for key,value in data.items():

                if 'SUBMITTAL' in key:
                    
                    for i,j in value.items():
                    
                        if "warranty" in i.lower() and j==[]:
                            sub_sec=key.split()[0]+"/"+i.split(" ", 1)[0].replace(".", "")
                            sub_section[sub_sec]=i.split(" ", 1)[1].strip()
                        
                        if j!=[]:
                            
                            for item in j:
                                if ("warranty" in item.lower()) or ("warranties" in item.lower()):
                                    sub_sec=key.split()[0]+"/"+i.split(" ", 1)[0].replace(".", "")
                                    sub_sec=sub_sec+"/"+item.split(" ", 1)[0].replace(".", "")
                                    sub_section[sub_sec]=item.split(" ", 1)[1].strip()
        if sub_section=={}:
            return None

        return sub_section

    def all(self,word,data):
        spec_num=Entries.spec_number(self)
        division=Entries.division(self,spec_num)
        spec_name=Entries.spec_name(self)
        sub_sections=Entries.search(self,word,data)

        data_dump=[]
        for s_no,(spec_sub_sec,sentence) in enumerate(sub_sections.items()):
            dump=[]
            dump.append(s_no+1)
            dump.append(division)
            dump.append(spec_num)
            dump.append(spec_name.title())
            dump.append(spec_sub_sec)

            if re.match('Manufacturer’s Warrant',sentence,re.I):
                dump.append(word+" - Manufacturer")
            elif re.match('Installer’s Warranty',sentence,re.IGNORECASE):
                dump.append(word+" - Installer")
            else:
                dump.append(word)

            dump.append(sentence)

            data_dump.append(dump)
        
        return data_dump

    def printcsv(self,data):
        # header = ['S.NO', 'Division', 'Spec Number','Spec Name','Spec Sub Section','Document Type', 'Description']
        with open('closeout.csv', 'a+', encoding='utf-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerows(data)

        f.close()


main=['01 73 29', '01 78 00', '01 78 36', '07 59 00', '07 62 00','07 72 00', '07 92 00', '08 11 13', '08 70 00', '08 80 00',
       '08 91 00', '09 51 00', '09 65 13', '09 69 00', '09 97 26','21 00 10', '21 01 00', '22 00 10', '22 01 00', '23 00 10',
       '23 05 13', '23 08 13', '23 09 00', '23 23 00', '23 36 00','23 74 13', '23 81 25', '23 84 13', '26 00 10', '26 05 00',
       '26 05 19', '26 05 26', '26 05 29', '26 05 33', '26 05 36','26 05 37', '26 05 38', '26 05 43', '26 05 48', '26 05 53',
       '26 08 13', '26 09 23', '26 22 00', '26 24 13i', '26 24 16','26 26 00i', '26 27 26', '26 28 13', '26 28 16', '26 32 13i',
       '26 33 53i', '26 41 13', '26 43 13', '26 51 00', '27 05 26','28 31 12']

def individual(spec,b):

    file='Docs/individual files/DLR(combined)/'+b[spec]
    get_data(file)

    list_data,str_data=read_data('raw_text.txt')
    parts=PartsDivision(str_data)
    headings=SectionHeading(parts)
    three_parts=parts_list(list_data)
    data=sub_section(three_parts,headings)

    word="Warranty"
    entry=Entries(list_data,str_data)
    # entry.search(word,data)
    entry.test_reports(data)
    # print(entry.spec_name())

def multiple(main,b,divs):

    header = ['S.NO', 'Division', 'Spec Number','Spec Name','Spec Sub Section','Document Type', 'Description']
    with open('closeout.csv', 'w+', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(header)
        f.close()
    zz=0
    for i in os.listdir(r'C:\Users\raghavg\vcons\Projects\submittals\Docs\individual files\DLR(combined)'):
        for num, pdf in b.items():
            if pdf==i:
                spec=num

        if (spec[:2] != '01'):
        # if spec in main:
            # print(spec[:2]=='01')
            print("============ Processing {} ==============".format(i))
            file='Docs/individual files/DLR(combined)/'+i
            try:
                get_data(file)
                list_data,str_data=read_data('raw_text.txt')
                parts=PartsDivision(str_data)
                headings=SectionHeading(parts)
                three_parts=parts_list(list_data)
                data=sub_section(three_parts,headings)

                word="Warranty"
                entry=Entries(list_data,str_data)
                ans=entry.search(word,data)
                if ans==None:
                    print("No warranty section found.")
                    zz=zz+1
                else:
                    entry.printcsv(entry.all(word,data))
                
            except Exception as e:
                print("Not able to process {}".format(i))
                zz=zz+1
    print(zz)

with open('read.pickle', 'rb') as handle:
    b = pickle.load(handle)

individual('03 30 00',b)

# divs=['01']
# divs=['07','08','09','21','22','23','26','27','28']

# multiple(main,b,divs)