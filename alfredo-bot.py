import subprocess, re, configparser, bs4 as bs, urllib.parse, pywinauto, pyautogui, cv2 as cv, os, sys, datetime, time
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from pywinauto.application import Application
from pywinauto import keyboard
from pywinauto import mouse

#/mnt/c/Users/Alex/Desktop/alfredo
if sys.platform == "win32":

	#Import the config.ini file
	config = configparser.ConfigParser()
	config.read('config.ini')

	#username, password, and teamname for alfresco
	alf_username = config['ALFRESCO']['alf_username']
	alf_password = config['ALFRESCO']['alf_username']
	alf_teamname = config['ALFRESCO']['alf_teamname']
	alf_sitename = config['ALFRESCO']['alf_sitename']

	#urlencode
	address = urllib.parse.quote(alf_teamname)
	#urlencode the '%' sign because that's how alfresco does it
	address = address.replace('%', '%25')

	#Import the current date
	local_dir = (os.getcwd() + "/IT395")
	currdate = "- " + str(datetime.datetime.today().strftime('%Y-%m-%d'))
	addr_date = str(datetime.datetime.today().strftime('%Y-%m-%d'))

	#Use the geckodriver
	browser = webdriver.Firefox()

	#create functions then execute them
	def kflux():
		"""
		Kill flux
		"""
		proc = subprocess.Popen(['taskkill','/IM','flux.exe'])
		proc.communicate()[0]
		proc.returncode
		if proc == 1:
			proc.terminate()
			sys.exit()

	def purge():
		"""
		Delete the IT395 folder in order to make sure there is new content being downloaded from gdrive
		"""
		proc = subprocess.Popen(['rclone.exe','purge','C:\\Users\\Alex\\Desktop\\alfredo\\IT395'])
		proc.communicate()[0]
		proc.returncode
		if proc == 1:
			proc.terminate()
			sys.exit()

	def gdrive(local):
		"""
		Kill flux
		Sync the IT395 folder to rclone using the terminal
		"""
		proc = subprocess.Popen(['rclone.exe','sync','gdrive:IT395', local])
		proc.communicate()[0]
		proc.returncode
		if proc == 1:
			proc.terminate()
			sys.exit()

	def alfresco():
		"""
		Automate Alfresco
		"""
		browser.set_window_position(512, 0)
		browser.set_window_size(1024, 768)
		browser.get("http://azuredrop.com")
		browser.find_element_by_id("page_x002e_components_x002e_slingshot-login_x0023_default-username").click()
		browser.find_element_by_id("page_x002e_components_x002e_slingshot-login_x0023_default-username").send_keys(alf_username)
		browser.find_element_by_id("page_x002e_components_x002e_slingshot-login_x0023_default-password").click()
		browser.find_element_by_id("page_x002e_components_x002e_slingshot-login_x0023_default-password").send_keys(alf_username)
		browser.find_element_by_id("page_x002e_components_x002e_slingshot-login_x0023_default-submit-button").click()
		browser.get("http://140.192.39.129:8080/share/page/site/" + alf_sitename + "/documentlibrary")
		browser.find_element_by_id("template_x002e_documentlist_v2_x002e_documentlibrary_x0023_default-createContent-button-button").click()
		browser.find_element_by_class_name("folder-file").click()
		time.sleep(3)
		browser.find_element_by_id("template_x002e_documentlist_v2_x002e_documentlibrary_x0023_default-createFolder_prop_cm_name").send_keys(alf_teamname + " " + currdate)
		browser.find_element_by_id("template_x002e_documentlist_v2_x002e_documentlibrary_x0023_default-createFolder_prop_cm_description").click()
		browser.find_element_by_id("template_x002e_documentlist_v2_x002e_documentlibrary_x0023_default-createFolder_prop_cm_description").send_keys("Created by " + alf_teamname + "'s alfredo-bot! On " + currdate)
		browser.find_element_by_id("template_x002e_documentlist_v2_x002e_documentlibrary_x0023_default-createFolder-form-submit-button").click()
		browser.get("http://140.192.39.129:8080/share/page/site/" + alf_sitename + "/documentlibrary#filter=path%7C%2F" + address + "%2520-%2520" + addr_date + "%7C&page=1")

	def windows():
		"""
		Automate Windows
		"""
		time.sleep(2)
		app = Application().start('explorer.exe')
		time.sleep(2)
		w_handle = pywinauto.findwindows.find_windows(title=u'File Explorer', class_name='CabinetWClass')[0]
		window = app.window(handle=w_handle)
		window.set_focus()
		ctrl = window['9']
		ctrl.set_focus()
		ctrl = window['9']
		ctrl.click_input()
		keyboard.SendKeys(os.getcwd() + "{ENTER} {SPACE} {DOWN} {DOWN}")
		if pyautogui.locateOnScreen(os.getcwd() + '\\Screenshots\\folder\\IT395folder2.PNG', confidence = .6) is not None:
			pyautogui.moveTo(pyautogui.locateCenterOnScreen(os.getcwd() + '\\Screenshots\\folder\\IT395folder2.PNG', confidence = .6))
			if pyautogui.locateOnScreen(os.getcwd() + '\\Screenshots\\website\\DROParea7.PNG', confidence = .6) is not None:
				pyautogui.dragTo(pyautogui.locateCenterOnScreen(os.getcwd() + '\\Screenshots\\website\\DROParea7.PNG', confidence = .6), duration=3.0, button='left')
			else:
				print("Could not locate drop area! Exiting...")
				sys.exit()
		else:
			print("Could not locate folder! Exiting...")
			sys.exit()

	def box(local):
		"""
		sync rclone folder to box.com
		"""
		print("Synced the local directory to box.com...")
		proc = subprocess.Popen(['rclone.exe','sync', local, 'box:IT395' + currdate])
		proc.communicate()[0]
		proc.returncode
		if proc == 1:
			proc.terminate()
			sys.exit()

	def flux():
		"""
		Turn flux back on
		"""
		proc = subprocess.Popen(['C:\\Users\\Alex\\AppData\\Local\\FluxSoftware\\Flux\\flux.exe'])
		proc.returncode
		if proc == 1:
			proc.terminate()
			sys.exit()
		print("Turned flux back on")
		
	kflux()
	gdrive(local_dir)
	alfresco()
	windows()
	box(local_dir)
	browser.close()
	flux()

else:
	#If Windows is not detected, then exit
	print("Win32 platform not detected! Exiting...")
	sys.exit()