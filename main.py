from io import BytesIO
from docx import Document
from fastapi import FastAPI, File, Request, UploadFile
import datetime
import os
from fastapi.responses import FileResponse, Response
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from docxtpl import DocxTemplate, RichText
from docxtpldvm import DVM

# pip install htmldocx
from htmldocx import HtmlToDocx

app = FastAPI()

app.add_middleware(
  CORSMiddleware,
  # allow_origins= ["http://172.16.200.180:9001"],
  allow_origins= ["*"],
  allow_credentials= True,
  allow_methods= ["*"],
  allow_headers= ["*"]
)

# @app.post("/items/")
# async def create_item(item: dict):
#     item_dict = item
#     print(item_dict)
#     if 'tax' in item_dict:
#         price_with_tax = item['price'] + item['tax']
#         item_dict.update({"price_with_tax": price_with_tax})
#     return item_dict

@app.post("/uploadfile/")
async def create_upload_file(file: UploadFile = File(...)):
    file_location = f"./templates/{file.filename}"
    with open(file_location, "wb") as f:
        f.write(await file.read())
    return {"info": f"file '{file.filename}' saved at '{file_location}'"}

@app.get("/get-ip/")
async def get_ip(request: Request):
    client_ip = request.client.host
    return {"client_ip": request.headers.get('X-Real-Ip')} # read nginx/default.conf

@app.get("/thbc/get-ip/")
async def get_ip(request: Request):
    client_ip = request.client.host
    return {"client_ip": client_ip} # for dev, test

from jinja2 import Environment, BaseLoader, Template

# Tạo hàm tùy chỉnh để chèn nội dung HTML vào tài liệu Word
def html_to_rich_text(html_content, _doc):
    desc_document = Document()
    new_parser = HtmlToDocx()
    new_parser.add_html_to_document(html_content , desc_document)
    buff = BytesIO()
    desc_document.save(buff)
    # buff.seek(0)
    return _doc.new_subdoc(buff)

def find_column_with_token(data, token_value):
    for item in data:
        if item.get('token') == token_value:
            return item.get('column')
    return None
def find_atr_with_token(data, token_value):
    for item in data:
        if item.get('token') == token_value:
            return item.get('atr')
    return None
def find_token_with_column(data, column_value):
    for item in data:
        if item.get('column') == column_value:
            return item.get('token')
    return None
def find_duplicate_indices(arr):
    indices = []
    i = 0
    while i < len(arr):
        start_index = i
        while i < len(arr) - 1 and arr[i] == arr[i + 1]:
            i += 1
        if start_index != i:
            indices.append([start_index+1, i+1])
        i += 1
    return indices
@app.post("/docx/")
async def gen_docx(data: dict):
    item= data['content']
    context= {}
    template = "./templates/"+ data['file'].replace('/', '').replace('\\', '')
    doc = DocxTemplate(template)
    for key in item:
        el= item[key]
        if 'html' in el and el['html'] is True:
            context.update({el['token']: html_to_rich_text(el['data'], doc)})
        else:
            if 'merges' in el:
                merges= el['merges']
                dat= el['data']
                if isinstance(dat, list):
                    tokens = [item["token"] for item in merges if "token" in item]
                    if len(tokens)> 0:
                        columns = [item["column"] for item in merges if "column" in item]
                        # print(tokens)
                        # print(columns)
                        prev_index= {}
                        if len(columns)== 0:
                            t= tokens[0]
                            prev_index.update({t: find_duplicate_indices(dat)})
                            # print(prev_index)
                            context.update({t: DVM(prev_index[t])})
                        else:
                            for t in tokens:
                                c= find_column_with_token(merges, t)
                                if c is not None:
                                    prev_index.update({t: find_duplicate_indices(
                                        [item[c] for item in dat if c in item]
                                    )})
                                    print(prev_index)
                                    context.update({t: DVM(prev_index[t])})
            elif 'merges_child' in el:
                merges= el['merges_child']
                dat= el['data']
                if isinstance(dat, list):
                    tokens = [item["token"] for item in merges if "token" in item]
                    atr = [item["atr"] for item in merges if "atr" in item]
                    columns = [item["column"] for item in merges if "column" in item]   
                    for t in tokens:
                        c= find_column_with_token(merges, t)
                        for d_0 in dat:
                            d= d_0[find_atr_with_token(merges, t)]
                            print('d:', d)
                            if c is not None:
                                vm_1= find_duplicate_indices(
                                    [item[c] for item in d if c in item]
                                )
                                d_0.update({t: DVM(vm_1)})
                                print('d_0', d_0)
            context.update({el['token']: el['data']})
    print(context)

    doc.render(context)
    # Tạo một buffer để lưu nội dung tài liệu
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    # Trả về nội dung tài liệu mà không cần lưu thành file
    return Response(
        buffer.read(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

    # # context.update({'diepnn': [(1,3),(4,6),(7,9)]})
    # return context
  



