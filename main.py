from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import scrapy
from scrapy.selector import Selector
from scrapy.crawler import CrawlerProcess
from openpyxl import Workbook
import time

class ClickSpider(scrapy.Spider):
    name = 'click_spider'
    
    def __init__(self, *args, **kwargs):
        super(ClickSpider, self).__init__(*args, **kwargs)
        
        self.visibility = expected_conditions.visibility_of_element_located
        service = Service(ChromeDriverManager().install())
        options = webdriver.ChromeOptions()
        options.add_argument("--window-size=1920,1080")
        self.driver = Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, 150)
        self.partidos = []
        self.start_date = None
        self.actualDay = ""

    def start_requests(self):
        urls = [
            'https://www.espn.co.cr/futbol/resultados'  # Reemplaza con la URL inicial
        ]
        for url in urls:
            yield scrapy.Request(
                url=url,
                callback=self.parse,
                headers={'unique-header': str(int(time.time()))}  # Añadir un header único
            )

    def parse(self, response):
        self.driver.get(response.url)
        self.driver.maximize_window()
        self.navigate_and_scrape()

        # Intentar hacer clic en el botón para seleccionar la siguiente date
        while True:
            try:
                DayContainer = self.wait.until(self.visibility((By.CSS_SELECTOR, "div.is-active")))
                """time.sleep(5)
                day = DayContainer.find_element(By.XPATH, './div/span[2]/span').text
                daySplit = day.split(" ")
                self.actualDay = daySplit[0]
                if self.start_date and self.actualDay == "1":
                    break"""

                next_div = DayContainer.find_element(By.XPATH, 'following-sibling::div')
                if not next_div.is_displayed():
                    break
                next_div.click()
                
                # Esperar a que la información se recargue y volver a extraer los datos
                self.wait.until(expected_conditions.staleness_of(DayContainer))
                self.navigate_and_scrape()
                
            except Exception as e:
                self.logger.info(f"Error al hacer clic en siguiente date o no hay más dates: {e}")
                break  # Salir del bucle si no hay más dates disponibles o ocurre un error
        self.save_to_excel()

    def navigate_and_scrape(self):
        selenium_response = Selector(text=self.driver.page_source)
        
        if not self.start_date:
            try:
                Calendar = self.wait.until(self.visibility((By.XPATH, "//*[@id='fittPageContainer']/div[2]/div[2]/div/div/div[1]/div/section/div/section/div/div/button")))
                Calendar.click()
                Day1 = self.wait.until(self.visibility((By.XPATH, "//*[@id='fittPageContainer']/div[2]/div[2]/div/div/div[1]/div/section/div/section/div/div/div[2]/div[2]/ul[1]/li[1]")))
                Day1.click()
                
                self.start_date = True
                return
            except Exception as e:
                self.logger.info(f"Error al abrir el calendario: {e}")

        date = selenium_response.css('h3::text').get()
        
        # Extrae información de los artículos en la página
        for sectionMain in selenium_response.css('section.Card.gameModules'):
            competicion = sectionMain.css('h3::text').get()
            for section in sectionMain.css('section.Scoreboard'):
                time = section.css('div.ScoreCell__Time::text').get()
                if(time != 'FT'):
                    continue
                
                teams = []
                for item in section.css('li'):
                    team = item.css('div.ScoreCell__TeamName::text').get()
                    score = item.css('div.ScoreCell__Score::text').get()
                    if(team != None and score != None):
                        teams.append({
                            'team': team,
                            'score' : score
                        })
                        
                partido = {}

                partido['date'] = date
                partido['competion'] = competicion
                i = 1
                for e in teams:
                    if(i%2!=0):
                        partido['local'] = e['team']
                        partido['localGoals'] = e['score']
                    else:
                        partido['visiting'] = e['team']
                        partido['visitingGoals'] = e['score']
                    i=i+1

                if(partido['local'] != ""):
                    self.partidos.append(partido)

    def save_to_excel(self):
        # Crear un libro de trabajo y una hoja
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = 'Matchs'

        # Escribir los encabezados
        headers = ['Date', 'Competion', 'Local Team', 'Visiting Team', 'Local Goals', 'Visiting Goals']
        sheet.append(headers)

        # Escribir los datos de los partidos
        for p in self.partidos:
            row = [
                p['date'],
                p['competion'],
                p['local'],
                p['visiting'],
                p['localGoals'],
                p['visitingGoals']
            ]
            sheet.append(row)

        # Guardar el archivo
        workbook.save('partidos.xlsx')

    def closed(self, reason):
        # Cierra el navegador de Selenium cuando el Spider termina
        self.save_to_excel()
        self.driver.quit()

# Ejecutar el spider
if __name__ == "__main__":
    process = CrawlerProcess(settings={
        'DUPEFILTER_CLASS': 'scrapy.dupefilters.BaseDupeFilter',
    })
    process.crawl(ClickSpider)
    process.start()