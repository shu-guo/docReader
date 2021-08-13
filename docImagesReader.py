import docx
import os
from docx.document import Document
from docx.parts.image import ImagePart

document = docx.Document('demo.docx')
# 创建imgPath
subImgPath = "./images"
if not os.path.exists(subImgPath):
    os.makedirs(subImgPath)
    print (subImgPath)

#print(document.part._rels)
rels = {}
for r in document.part.rels.values():
    if isinstance(r._target, docx.parts.image.ImagePart):
        rels[r.rId] = os.path.basename(r._target.partname)
        print ('rels:', r.rId )
        imgName = r.rId + '.png'
        with open(subImgPath + "/" + imgName,"wb") as f:
        	f.write(r.target_part.blob)

'''
index = 0
for rel in document.part._rels:
    rel = document.part._rels[rel]               #获得资源
    print (rel.target_ref)
    if "image" not in rel.target_ref:
        continue
    index=''
    for r in document.part.rels.values():
    	if isinstance(r._target, docx.parts.image.ImagePart):
	        #rels[r.rId] = os.path.basename(r._target.partname)
	        index=r.rId
	        print ('index:', r.rId )
    imgName = index + '.png'
    with open(subImgPath + "/" + imgName,"wb") as f:
        f.write(rel.target_part.blob)

'''