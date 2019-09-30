#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import sys
import tempfile
import pdf2image
import cv2
import numpy as np
from PIL import Image
import pytesseract
import os
import imghdr
from fpdf import FPDF
import img2pdf
from PyPDF2 import PdfFileMerger
import shutil

private_file = pd.read_excel('private.xlsx')
path_to_out_files_dir = private_file.iloc[0][1]
path_to_final_pdf_files_dir = private_file.iloc[1][1]
path_to_original_pdf_file = private_file.iloc[2][1]
path_to_out_dir = private_file.iloc[3][1]

def removeAccents(input_text):
    strange='ŮôῡΒძěἊἦëĐᾇόἶἧзвŅῑἼźἓŉἐÿἈΌἢὶЁϋυŕŽŎŃğûλВὦėἜŤŨîᾪĝžἙâᾣÚκὔჯᾏᾢĠфĞὝŲŊŁČῐЙῤŌὭŏყἀхῦЧĎὍОуνἱῺèᾒῘᾘὨШūლἚύсÁóĒἍŷöὄЗὤἥბĔõὅῥŋБщἝξĢюᾫაπჟῸდΓÕűřἅгἰშΨńģὌΥÒᾬÏἴქὀῖὣᾙῶŠὟὁἵÖἕΕῨčᾈķЭτἻůᾕἫжΩᾶŇᾁἣჩαἄἹΖеУŹἃἠᾞåᾄГΠКíōĪὮϊὂᾱიżŦИὙἮὖÛĮἳφᾖἋΎΰῩŚἷРῈĲἁéὃσňİΙῠΚĸὛΪᾝᾯψÄᾭêὠÀღЫĩĈμΆᾌἨÑἑïოĵÃŒŸζჭᾼőΣŻçųøΤΑËņĭῙŘАдὗპŰἤცᾓήἯΐÎეὊὼΘЖᾜὢĚἩħĂыῳὧďТΗἺĬὰὡὬὫÇЩᾧñῢĻᾅÆßшδòÂчῌᾃΉᾑΦÍīМƒÜἒĴἿťᾴĶÊΊȘῃΟúχΔὋŴćŔῴῆЦЮΝΛῪŢὯнῬũãáἽĕᾗნᾳἆᾥйᾡὒსᾎĆрĀüСὕÅýფᾺῲšŵкἎἇὑЛვёἂΏθĘэᾋΧĉᾐĤὐὴιăąäὺÈФĺῇἘſგŜæῼῄĊἏØÉПяწДĿᾮἭĜХῂᾦωთĦлðὩზკίᾂᾆἪпἸиᾠώᾀŪāоÙἉἾρаđἌΞļÔβĖÝᾔĨНŀęᾤÓцЕĽŞὈÞუтΈέıàᾍἛśìŶŬȚĳῧῊᾟάεŖᾨᾉςΡმᾊᾸįᾚὥηᾛġÐὓłγľмþᾹἲἔбċῗჰხοἬŗŐἡὲῷῚΫŭᾩὸùᾷĹēრЯĄὉὪῒᾲΜᾰÌœĥტ'
    ascii_replacements='UoyBdeAieDaoiiZVNiIzeneyAOiiEyyrZONgulVoeETUiOgzEaoUkyjAoGFGYUNLCiIrOOoqaKyCDOOUniOeiIIOSulEySAoEAyooZoibEoornBSEkGYOapzOdGOuraGisPngOYOOIikoioIoSYoiOeEYcAkEtIuiIZOaNaicaaIZEUZaiIaaGPKioIOioaizTIYIyUIifiAYyYSiREIaeosnIIyKkYIIOpAOeoAgYiCmAAINeiojAOYzcAoSZcuoTAEniIRADypUitiiIiIeOoTZIoEIhAYoodTIIIaoOOCSonyKaAsSdoACIaIiFIiMfUeJItaKEISiOuxDOWcRoiTYNLYTONRuaaIeinaaoIoysACRAuSyAypAoswKAayLvEaOtEEAXciHyiiaaayEFliEsgSaOiCAOEPYtDKOIGKiootHLdOzkiaaIPIIooaUaOUAIrAdAKlObEYiINleoOTEKSOTuTEeiaAEsiYUTiyIIaeROAsRmAAiIoiIgDylglMtAieBcihkoIrOieoIYuOouaKerYAOOiaMaIoht'
    translator=str.maketrans(strange,ascii_replacements)
    return input_text.translate(translator)

def find_bottommost(img):
	gray = cv2.cvtColor(img,cv2.COLOR_BGR2GRAY)
	thresh = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,11,2) #make in binary image for threshold value of 1 (for better accuracy)
	thresh_inv = cv2.bitwise_not(thresh)
	img2,contours,hierarchy = cv2.findContours(thresh_inv, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE) #thresh is our source image, 2nd argument is contour retrieval mode (hierarchy between contours) and 3rd is contour approximation method
	cnt = contours[0]
	bottommost = tuple(cnt[cnt[:,:,1].argmax()][0])
	return img[0:bottommost[1]]

def crop_image_borders(image_path):
	img = cv2.imread(image_path) #imread reads image (syntax: path/name, way image should be read - color or 1 (default), grayscale or 0, unchanged or -1)
	height1, width1, channels = img.shape
	crop = img[int(5*height1/100):height1-int(5*height1/100),0:width1] #cuts image for above coordinates from array
	crop2 = find_bottommost(crop)
	height2, width2, channels = crop2.shape
	crop3 = crop2[0:height2,int(width2/2):width2]
	crop4 = find_bottommost(crop3)
	height3, width3, channels = crop4.shape
	crop5 = crop2[0:height3,0:width2]
	height4, width4, channels = crop5.shape
	surname_scale_short = 93
	surname_scale_long = 96
	if height4 < 1300:
		crop_surname = Image.fromarray(crop5[0:height4-int(surname_scale_short*height4/100),0:width4-int(60*width4/100)])
		surname = removeAccents(pytesseract.image_to_string(crop_surname, lang='pol').split(": ")[1])
		if not os.path.exists('C:\\Users\\Paulina\\Desktop\\Python\\7. Ucinanie pdf\\out files\\'+surname):
			os.makedirs('C:\\Users\\Paulina\\Desktop\\Python\\7. Ucinanie pdf\\out files\\'+surname)
		cv2.imwrite('C:\\Users\\Paulina\\Desktop\\Python\\7. Ucinanie pdf\\out files\\'+surname+'\\'+os.path.basename(image_path),crop5) #imwrite saves image (syntax: file name, image you want to save)
	else:
		crop_surname_long = Image.fromarray(crop5[0:height4-int(surname_scale_long*height4/100),0:width4-int(60*width4/100)])
		surname_long = removeAccents(pytesseract.image_to_string(crop_surname_long, lang='pol').split(": ")[1])
		if not os.path.exists('C:\\Users\\Paulina\\Desktop\\Python\\7. Ucinanie pdf\\out files\\'+surname_long):
			os.makedirs('C:\\Users\\Paulina\\Desktop\\Python\\7. Ucinanie pdf\\out files\\'+surname_long)
		cv2.imwrite('C:\\Users\\Paulina\\Desktop\\Python\\7. Ucinanie pdf\\out files\\'+surname_long+'\\'+os.path.basename(image_path),crop5) #imwrite saves image (syntax: file name, image you want to save)

def merge_images_to_pdf(path_to_folder):
	#page_height = 3508 #height of A4 page
	page_height = 3508 #height of A4 page
	file_list = []
	for image in os.listdir(path_to_folder):
		if "jpg" in image.split(".")[-1]:
			file_path = path_to_folder+"\\"+image
			img = cv2.imread(file_path)
			height, width, channels = img.shape
			file_list.append([file_path, height])

	number_of_full_page = 1
	number_of_pdf = 1
	whole_pdf_file = FPDF()
	while(len(file_list)):
		actuall_len = 0
		single_page = np.zeros((1,2480,3), np.uint8)
		while( len(file_list) and ( (actuall_len + file_list[0][1]) < page_height) ):
			actual_image = cv2.imread(file_list[0][0])
			single_page = np.vstack((single_page, actual_image))
			actuall_len += file_list[0][1]
			file_list = file_list[1:]
		tmp_img = path_to_folder+"\\"+str(number_of_full_page)+"tmp.jpg"
		cv2.imwrite(tmp_img, single_page)
		number_of_full_page += number_of_full_page

		#creating empty pdf file
		pdf_bytes = img2pdf.convert([tmp_img])
		os.remove(tmp_img)
		whole_pdf_file.compress = False
		whole_pdf_file.add_page(orientation="Portrait")
		whole_pdf_file.output(path_to_folder+"\\"+os.path.basename(path_to_folder)+str(number_of_pdf)+".pdf", "F")

		#fill above empty pdf with image
		file = open(path_to_folder+"\\"+os.path.basename(path_to_folder)+str(number_of_pdf)+".pdf","wb")
		file.write(pdf_bytes)
		number_of_pdf += number_of_pdf
		file.close()


def final_pdf(path_to_folder):
	#combining pdf files for each surname folder
	merger = PdfFileMerger()
	files_list = []
	pdf_list = []
	files_list = os.listdir(path_to_folder)
	pdf_list = [f for f in files_list if f.endswith('.pdf')]

	for file in pdf_list:
		merger.append(path_to_folder+"\\"+file)
	merger.write(path_to_folder+"\\"+os.path.basename(path_to_folder)+".pdf")
	merger.close()
	for file in files_list:
		if not file.endswith('.pdf'):
			os.remove(path_to_folder+"\\"+file)

def move_to_folder(path_to_folder3):
	folder_list = []
	for folder in os.listdir(path_to_folder3):
		for file6 in os.listdir(path_to_folder3+folder+'\\'):
			shutil.move(path_to_folder3+folder+'\\'+file6, path_to_final_pdf_files_dir)

def remove_useless_files(path_to_folder):
	number_list = list(range(10))
	for file in os.listdir(path_to_folder):
		if file[-5:-4].isdigit():
			os.remove(path_to_folder+'\\'+file)

def main():
	with tempfile.TemporaryDirectory() as path:
		images_from_path = pdf2image.convert_from_path(pdf_path = path_to_original_pdf_file, dpi=300, output_folder = path_to_out_dir, first_page=None, last_page=None, fmt='jpg', thread_count=1, userpw=None)

		for idx in range(len(images_from_path)):
			crop_image_borders(images_from_path[idx].filename)

		for folder in os.listdir(path_to_out_files_dir):
			merge_images_to_pdf(path_to_out_files_dir+folder)
		for folder in os.listdir(path_to_out_files_dir):
			final_pdf(path_to_out_files_dir + folder)
		if not os.path.exists(path_to_final_pdf_files_dir):
			os.makedirs(path_to_final_pdf_files_dir)
		move_to_folder(path_to_out_files_dir)
		remove_useless_files(path_to_final_pdf_files_dir)

if __name__ == "__main__":

	main()
