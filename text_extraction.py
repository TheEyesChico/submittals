import pdfplumber

def get_data(path,start=None,end=None):

    pdf = pdfplumber.open(path)

    if start==None and end==None:
        docs=pdf.pages
    elif end==None:
        docs=pdf.pages[start-1:start]
    else:
        docs=pdf.pages[start-1:end]
    
    with open('raw_text.txt', 'w',encoding="utf-8") as f:
        for pdf_page in docs:
            roi = pdf_page.within_bbox((0, 62, pdf_page.width, 761))   
            # previous calls : 2nd - 67, 4th - 763
            f.write(roi.extract_text())

    f.close()
    


