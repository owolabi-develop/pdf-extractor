import fitz
from pprint import pprint
import pdfplumber
import pandas as pd
import mimetypes
import os,tempfile
import pymongo
from pymongo.mongo_client import MongoClient
import docx
import io

from elasticsearch import Elasticsearch


def get_table(filepath):
   with pdfplumber.open(filepath) as pdf:
    # iterate over each page
    for page in pdf.pages:
        tables = page.extract_tables()
        return tables
    

def image_data(filepath):
   images_data = None
   image = fitz.open(filepath)
   for page_num in range(len(image)):
      page_content = image[page_num]
      if len(page_content.get_images()) <=0:
         return images_data
      else:
       ## get image bytes
       images_data = image.extract_image(page_content.get_images()[0][0])
       return images_data

def get_links(filepath):
    links_data = None
    doc = fitz.open(filepath)
    for index in range(len(doc)):
        pages = doc.load_page(index)
        if len(pages.get_links())<=0:
           links_data
        else:
          links_data = pages.get_links()[0]['uri'] 
          return links_data 



def get_all_content(filepath):
    page_data = {"pages":[]}

    doc = fitz.open(filepath)

    for index in range(len(doc)):
        tables = get_table(filepath) 
        image = image_data(filepath)
        pages = doc.load_page(index)
        links = get_links(filepath)

        
        text_data = {
            "page":pages.number,
            "text":pages.get_text(),
            "links":links, 
            "table": tables,
            "images": f"{image}"  
        }
        page_data['metadata'] = doc.metadata,
        page_data["pages"].append(text_data)

    return page_data
    

def get_flat_data(file):

    #csv = .xls
    #exce file =.xlsx
    file_type = mimetypes.guess_type(file)[0]
    file_extension = mimetypes.guess_extension(file_type)

    if file_extension == ".xls":
        df =  pd.read_csv(file,dtype='object')
        #print(df.head())

        df.drop_duplicates(inplace=True)
        df.loc[:] = df.loc[:].apply(lambda x:x.str.lower())\
        .fillna("None")
        return df.to_dict()
        
    elif file_extension == ".xlsx":
        df =  pd.read_excel(file,engine="openpyxl",dtype='str')
        
        df.drop_duplicates(inplace=True)
        df.loc[:] = df.loc[:].apply(lambda x:x.str.lower())\
        .fillna("None")
        return df.to_dict()
           


def extract_from_word(filepath):
   docx_dict = {}
   document = docx.Document(filepath)
   index = 0
   for para in document.paragraphs:
      index+=1
      if (len(para.text))>0:
         docx_dict[index] = para.text

   return docx_dict
         


# def upload_data_mongodb(data=None):
#    client = MongoClient("mongodb+srv://owolabi:84563320owo@scrapper.joaxnnt.mongodb.net/?retryWrites=true&w=majority")
#    db = client['pdfdata']
#    collection = db['pdfdata']
#    collection.insert_one(data)




def upload_data_to_elasticsearch(doc):
    client = Elasticsearch(
       cloud_id='',
       basic_auth=("","")
    )
    resp= client.index(index="pdf_data",document=doc)
    print(resp['result'])




#content = get_all_content('sample2.pdf')

#upload_data_to_elasticsearch(content)

#flat_data = get_flat_data('Financial.xlsx')
    
#upload_data_to_elasticsearch(flat_data)

#pprint(data)



