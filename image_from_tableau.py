import tableauserverclient as TSC
import pandas as pd
import PyPDF2 as pf
import time
from datetime import date
import os
import pandas.io.sql as psql
from PIL import Image, ImageOps

today = date.today()
date_today = today.strftime("%Y-%m-%d")

# Sign in to Tableau Server
tableau_auth = TSC.TableauAuth('****', '*****', site_id='')
server = TSC.Server('https://***********')
server.auth.sign_in(tableau_auth)
server.version = "3.4"
current_path = os.getcwd()
# Get the workbook and view
#workbook_name = 'Pricing Dynamics - Comercial Castro'
#view_name = 'Alerta Seguimiento Precios propios'

# workbook = server.views.get_by_id('def557f0-4dd9-449a-8555-e46e646a7a34')
# get the view information
#server.workbooks.populate_views(workbook)
req_option = TSC.RequestOptions()

view = server.views.get_by_id('***********************')
print(view.name)

req_option.filter.add(TSC.Filter(TSC.RequestOptions.Field.Name, TSC.RequestOptions.Operator.Equals, view.name))
all_views, pagination_item = server.views.get(req_option)
if not all_views:
        raise LookupError("No se encontro la vista: "+view)
view_item = None
view_item = all_views[0]

image_req_option = TSC.ImageRequestOptions(imageresolution=TSC.ImageRequestOptions.Resolution.High)
server.views.populate_image(view_item, image_req_option)  


with open(current_path+"/send_files/"+view.name+" "+date_today+".jpg", "wb") as image_file:
        image_file.write(view_item.image)
image1 = ImageOps.expand(Image.open(current_path+"/send_files/"+view.name+" "+date_today+".jpg"), border = 50, fill = 'white')
im1 = image1.convert('RGB')
im1.save(current_path+"/send_files/"+view.name+" "+date_today+".pdf",resolution=100.0, save_all=True)


print('save')

# Sign out from Tableau Server
#server.auth.sign_out()
