from cgitb import reset
from scipy.spatial.distance import euclidean

from imutils import perspective
import openpyxl

from imutils import contours
import pandas as pd

import numpy as np

import imutils

import cv2
import ast


# To show array of images 

def show_images(images):

 for i, img in enumerate(images):

  cv2.imshow("image_" + str(i), img)

 cv2.waitKey(0)

 cv2.destroyAllWindows()



height = []
width = []
areas = []
final_perimeter = []

#File path or image to be analysed
img_path = r"C:\Users\olumi\OneDrive\Desktop\AGE 417 Grain Samples\Sampea 10\Sampea 10 c.jpeg.jpeg"

# Read image and preprocess

image = cv2.imread(img_path)

gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

blur = cv2.GaussianBlur(gray, (9, 9), 0)

edged = cv2.Canny(blur, 50, 100)

edged = cv2.dilate(edged, None, iterations=1)

edged = cv2.erode(edged, None, iterations=1)


#show_images([blur, edged])
cv2.imshow('Grayscale', gray)
# cv2.imshow('blur', blur)
cv2.imshow('Binary', edged)

# Find contours

cnts = cv2.findContours(edged.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

cnts = imutils.grab_contours(cnts)

# Sort contours from left to right as leftmost contour is reference object

(cnts, _) = contours.sort_contours(cnts)

# Remove contours which are not large enough

cnts = [x for x in cnts if cv2.contourArea(x) > 100]

#cv2.drawContours(image, cnts, -1, (0,255,0), 3)



number = (len(cnts))
print ("There are " + str(number) + " items in the image")

ref_object = cnts[0]

box = cv2.minAreaRect(ref_object)

box = cv2.boxPoints(box)

box = np.array(box, dtype="int")

box = perspective.order_points(box)

(tl, tr, br, bl) = box

dist_in_pixel = euclidean(tl, tr)

dist_in_cm = 0.8

pixel_per_cm = dist_in_pixel/dist_in_cm

# Draw remaining contours

for cnt in cnts:

 box = cv2.minAreaRect(cnt)

 box = cv2.boxPoints(box)

 box = np.array(box, dtype="int")

 box = perspective.order_points(box)

 (tl, tr, br, bl) = box

 cv2.drawContours(image, [box.astype("int")], -1, (0, 0, 255), 2)

 mid_pt_horizontal = (tl[0] + int(abs(tr[0] - tl[0])/2), tl[1] + int(abs(tr[1] - tl[1])/2))

 mid_pt_verticle = (tr[0] + int(abs(tr[0] - br[0])/2), tr[1] + int(abs(tr[1] - br[1])/2))

 wid = euclidean(tl, tr)/pixel_per_cm
 wid = round(wid, 2)

 ht = euclidean(tr, br)/pixel_per_cm
 ht = round(ht, 2)



#print result to excel spreadsheet


 my_height = []
 
 height.append(ht)

 width.append(wid)
 

 count = 0
 area = wid * ht
 
 area = round(area, 2)
 areas.append(area)
 
 


 perimeter = (2*(ht + wid))
 perimeter = round(perimeter, 2)
 final_perimeter.append(perimeter)


#  print(perimeter)
  #print(f"length = {i}cm, width = {j}cm, area = {k}cm")
 

#  dict = {'height': i, 'width': j, 'area':k} 
#  df = pd.DataFrame(dict)
#  a = len(dict)
#  print(dict)
#  print(len(str(i)))
#  df.to_csv('Samuel.csv')
#  df.to_excel(r'C:\Users\User\count\binary.xlsx', index = False)


 
 cv2.putText(image, "{:.1f}cm".format(wid), (int(mid_pt_horizontal[0] - 15), int(mid_pt_horizontal[1] - 10)), 

  cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 0), 2)

 cv2.putText(image, "{:.1f}cm".format(ht), (int(mid_pt_verticle[0] + 10), int(mid_pt_verticle[1])), 

  cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 0), 2)

show_images([image])


# df.to_excel(r'C:\Users\User\Desktop\417\Solution1.xlsx', index = False)
 
heightdf = pd.DataFrame(columns=['height'])
heightdf['height'] = height
widthdf = pd.DataFrame(columns=['width'])
widthdf['width'] = width
Areadf = pd.DataFrame(columns=['area'])
Areadf['area'] = areas
perimeterdf = pd.DataFrame(columns=['final_perimeter'])
perimeterdf['final_perimeter'] = final_perimeter

#writer = heightdf.merge(widthdf,how ='left').merge(Areadf,how ='left').merge(perimeterdf,how ='left')
writer = pd.concat([heightdf, widthdf, Areadf, perimeterdf], axis=1, join="inner")

# creating excel writer object


writer.to_excel(r'C:\Users\olumi\Downloads\project works\age417prac\sampea10c.xlsx')
 
# save the excel
print("codecompleted successfully")