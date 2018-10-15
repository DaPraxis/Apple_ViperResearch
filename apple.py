import requests
from bs4 import BeautifulSoup as soup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import numpy as np
import pandas as pd
import xlsxwriter
import datetime
import math

def init(myurl):
	driver = webdriver.Chrome(executable_path='')
	driver.get(myurl)
	return driver

def findInfo(driver):
	Info = []
	xpath1 = '//*[@id="tabs_dimensionCapacity"]/fieldset/ul/li[1]/div[2]/div[1]/div/div[1]/span[2]'
	xpath2 = '//*[@id="tabs_dimensionCapacity"]/fieldset/ul/li[2]/div[2]/div[1]/div/div[1]/span[2]'
	xpath3 = '//*[@id="tabs_dimensionCapacity"]/fieldset/ul/li[3]/div[2]/div[1]/div/div[1]/span[2]'
	xpath = {xpath1,xpath2,xpath3}
	for x in xpath:
		element = WebDriverWait(driver, 10).until(
        	EC.presence_of_element_located((By.XPATH, x))
    	)
		content = driver.find_elements_by_xpath(x)
		for a in content:
			Info.append(a.text)
	return Info

def export_to_excel(table):
	writer = pd.ExcelWriter('',engine='xlsxwriter')
	Workbook=writer.book
	table.to_excel(writer,sheet_name='Sheet1')
	worksheet=writer.sheets['Sheet1']
	writer.save()
	# writer = pd.ExcelWriter('Apple.xlsx')
	# table.to_excel(writer,'Apple1')
	# writer.save()
#def export_to_csv(table):
#	table.to_csv('D:\\Python\\Projects_for_Viper\\Apple.csv', sep=',')

def read_excel():
	df_out=pd.read_excel(open('','rb'))
	#with pd.option_context('display.max_rows', None, 'display.max_columns', None):
	#    print(df)
	url_line = df_out.loc[df_out['Unnamed: 1'] =='Source']
	url_line = url_line.iloc[0].tolist()
	URL=[]
	for a in url_line:
		a=str(a)
		if a!='nan' and a!='Source':
			URL.append(a)
	return URL

def region_Change(url):
	a=url.split('ttps://www.apple.com')[1]
	region=a.split('shop/buy-iphone/')[0]
	if region == '/':
		region = 'us'
	else:
		region = region[1:-1]
	return region

def generate_final_1(urls):
	Final = []
	for url in urls:
		time = datetime.datetime.now()
		content = [time,url]
		driver = init(url)
		content.append(region_Change(url))
		content.extend(findInfo(driver))
		Final.append(content)
	col = ['Time','URL','Region','64GB','256GB','512GB']
	sheet = pd.DataFrame(Final,columns=col)
	return sheet
	
def generate_final_2(urls):
	col_num = len(urls)
	final=[]
	start = datetime.datetime.now()
	final.append(datetime.datetime.now().time())
	final.append(datetime.datetime.now().date())
	print('in total of %d website needed to be processed'%col_num)
	counter=1	
	for url in urls:
		print('processing %d website'% counter)
		driver = init(url)
		if counter>=10:
			final.append(findInfo(driver)[2])
		else:
			final.append(findInfo(driver)[1])
		print('%d done'%(counter))
		counter=counter+1
		driver.close()
	print('updated parts:')
	print (final)
	end=datetime.datetime.now()
	print('elapse: %s'%str(end-start))
	#sheet = pd.DataFrame(final).T
	return final

def appendMax(df, part):
	height=df.shape[0]
	length=df.shape[1]
	len_part=len(part)
	for i in range(length-len_part):
		part.append('')
	df.loc[height]=part
	return df
	
	
	
if __name__ == '__main__':
	urls=read_excel()
	sheet = generate_final_2(urls)
	df_out=appendMax(df_out,sheet)
	#df_out.append(sheet, ignore_index=True)
	export_to_excel(df_out)
#	export_to_csv(df_out)

	