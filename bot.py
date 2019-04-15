"""
Choose your settings in the GUI and then let the bot add every person
corresponding to a keyword entered in the LinkedIn's search bar.
"""
import os
import sys

# For executable
if getattr(sys, 'frozen', False):
	basePath = sys._MEIPASS
else:
	basePath = os.path.dirname(os.path.abspath(__file__))

import time
import tkinter as tk
from tkinter import messagebox

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, WebDriverException

WELCOME_MESSAGE = "Message joined to the Linkedin's contact request"
JOB_DONE = "http://www.jobdoneremovals.co.uk/assets/images/job-done-logo-1734x552.png"
LOGINED_URL = "https://www.linkedin.com/feed/"
ICON = "linkedin.ico"


def showError(title, msg):
	"""
	Show an error message box displaying msg.
	"""
	temp = tk.Tk()
	temp.withdraw()
	messagebox.showerror(title, msg)
	temp.destroy()


class LinkeBot(object):
	"""
	Class managing the bot.
	"""
	base = "https://fr.linkedin.com/"
	resultClassName = "search-result__occluded-item"
	maxNResultPages = 100

	def __init__(self):
		self._profileInit()

		try:
			self.driver = webdriver.Firefox(executable_path=os.path.join(basePath, "geckodriver.exe"),
											firefox_profile=self.profile, log_path='NUL')
		except WebDriverException:
			showError("Erreur", "Ce programme nécessite Mozilla Firefox pour fonctionner.")

		self.driver.maximize_window()

	def _profileInit(self):
		self.profile = webdriver.FirefoxProfile()
		self.profile.set_preference("browser.privatebrowsing.autostart", True)

	def close(self):
		self.driver.close()

	def quit(self):
		self.driver.quit()

	def go(self, url):
		self.driver.get(url)

	def makeSearchUrl(self, keyword):
		return "https://www.linkedin.com/search/results/people/?keywords={}&origin=GLOBAL_SEARCH_HEADER".format(keyword)

	def byClass(self, name, where=None):
		where = self.driver if where is None else where
		return where.find_element_by_class_name(name)

	def byClasses(self, name, where=None):
		where = self.driver if where is None else where
		return self.driver.find_elements_by_class_name(name)

	def byXpath(self, xpath):
		return self.driver.find_element_by_xpath(xpath)

	def getButtonFromLi(self, li):
		try:
			b = li.find_element_by_tag_name("button")
		except NoSuchElementException:
			return None

		if b.text != "Se connecter":
			return None

		return b

	def getAllButtonsByResultPage(self):
		return [self.getButtonFromLi(li) for li in self.byClasses(self.resultClassName) if
				self.getButtonFromLi(li) is not None]

	def writeNote(self, msg):
		div = self.byClass("send-invite__actions")
		self.byClass("button-secondary-large", where=div).click()
		self.byXpath('//*[@id="custom-message"]').send_keys(msg)
		self.byClass("button-primary-large", where=div).click()

	def scroll(self, pauseTime=0.02, pixels=10, iters=-1, endOfPage=False):
		current = 0
		i = 0

		maxPixel = self.driver.execute_script("return document.body.scrollHeight") if endOfPage else -1

		while True:
			current += pixels
			i += 1

			self.driver.execute_script("window.scrollTo({}, {});".format(current - pixels, current))
			time.sleep(pauseTime)

			if iters > 0 and i >= iters:
				break

			if maxPixel > 0 and current >= maxPixel:
				break

	def nextResultPage(self, keyword):
		i = 1
		while True:
			if i > self.maxNResultPages:
				break

			yield self.makeSearchUrl(keyword) + "&page=" + str(i)
			i += 1

	def login(self, login, password):
		self.driver.get(self.base)

		mailE = self.byXpath('//*[@id="login-email"]')
		passwordE = self.byXpath('//*[@id="login-password"]')
		submitE = self.byXpath('//*[@id="login-submit"]')

		mailE.send_keys(login)
		passwordE.send_keys(password)
		submitE.click()

		if self.driver.current_url != LOGINED_URL:
			self.quit()
			raise ValueError("Mauvaises informations de connexion à Linkedin")

	def main(self, mail, password, startMessage, keyword, pages):
		pagesGene = self.nextResultPage(keyword)

		self.login(mail, password)

		for i in range(pages):
			try:
				self.go(next(pagesGene))
			except StopIteration:
				break

			self.scroll(endOfPage=True)

			for b in self.getAllButtonsByResultPage():
				good = False
				count = 0
				while not good:

					try:
						b.click()
						self.writeNote(startMessage)
						break

					except ElementClickInterceptedException:
						time.sleep(0.5)
						count += 1
						if count == 8:
							break

						continue

		self.go(JOB_DONE)
		time.sleep(5)

		self.quit()

		return True


class GUI(object):
	"""
	Main class for the TkInter GUI.
	"""

	def __init__(self):
		self.root = tk.Tk()
		self.customWindow()
		self.setAll()
		self.run()

	def customWindow(self):
		self.root.title("Linkebot")
		self.root.iconbitmap(os.path.join(basePath, ICON))

		self.root.protocol("WM_DELETE_WINDOW", sys.exit)

		w = 700
		h = 395

		x = (self.root.winfo_screenwidth() / 2) - (w / 2)
		y = (self.root.winfo_screenheight() / 2) - (h / 2)

		self.root.geometry("{}x{}+{}+{}".format(w, h, int(x), int(y)))
		self.root.resizable(0, 0)

	def _hide(self):
		self.root.withdraw()

	def quit(self):
		self.startMessage = self.textE.get('1.0', tk.END)
		self.root.destroy()

	def setAll(self):
		self.root.columnconfigure(2, minsize=115)
		self.root.columnconfigure(3, minsize=115)

		self.var1 = tk.StringVar()
		self.var2 = tk.IntVar()
		self.login = tk.StringVar()
		self.password = tk.StringVar()

		tk.Label(self.root, text="Configuration du bot", font="Helvetica 12 bold").grid(row=0, columnspan=4, pady=15)

		tk.Label(self.root, text="Mots à rechercher :", width=17, anchor="w", font="Helvetica 10").grid(row=1, column=0,
																										padx=3)
		tk.Label(self.root, text="Pages à traiter :", width=17, anchor="w", font="Helvetica 10").grid(row=2, column=0,
																									  padx=3)
		tk.Label(self.root, text="Message connection :", width=17, anchor="nw", font="Helvetica 10").grid(row=3,
																										  column=0,
																										  padx=3)

		tk.Label(self.root, text="Login Linkedin :", width=17, anchor="nw", font="Helvetica 10").grid(row=4, column=0,
																									  padx=3)
		tk.Label(self.root, text="Mdp Linkedin :", width=17, anchor="nw", font="Helvetica 10").grid(row=5, column=0,
																									padx=3)

		self.keywordE = tk.Entry(self.root, textvariable=self.var1)
		self.pagesE = tk.Scale(self.root, variable=self.var2, from_=0, to=100, resolution=1, tickinterval=20,
							   orient="horizontal")
		self.textE = tk.Text(self.root, height=8, width=67)

		self.loginE = tk.Entry(self.root, textvariable=self.login)
		self.passwordE = tk.Entry(self.root, textvariable=self.password, show="*")

		self.keywordE.grid(row=1, column=2, columnspan=2, sticky='nswe')
		self.pagesE.grid(row=2, column=2, columnspan=2, sticky='nswe')
		self.textE.grid(row=3, column=2, columnspan=2, pady=8)

		self.loginE.grid(row=4, column=2, columnspan=2, sticky='nswe', pady=8)
		self.passwordE.grid(row=5, column=2, columnspan=2, sticky='nswe', pady=8)

		self.textE.insert(tk.END, WELCOME_MESSAGE)

		tk.Button(self.root, text="Quitter", width=15, command=sys.exit).grid(row=6, column=0, pady=10, columnspan=2,
																			  sticky="e")
		tk.Button(self.root, text="Valider", width=15, command=self.quit).grid(row=6, column=2, pady=10, columnspan=2,
																			   sticky="e")

	def run(self):
		tk.mainloop()

	def getVar1(self):
		return self.var1.get()

	def getVar2(self):
		return self.var2.get()

	def getText(self):
		return self.startMessage

	def getLogin(self):
		return self.login.get()

	def getPassword(self):
		return self.password.get()


if __name__ == "__main__":
	"""
	Main part of the program.
	"""
	g = GUI()

	keyword = g.getVar1()
	nPages = g.getVar2()
	startMessage = g.getText()
	login = g.getLogin()
	password = g.getPassword()

	l = LinkeBot()

	try:
		l.main(login, password, startMessage, keyword, nPages)

	except ValueError:
		showError("Erreur", "Mauvaises informations de connexion à Linkedin")

	sys.exit(0)
