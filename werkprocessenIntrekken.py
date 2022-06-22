# %%
import pandas as pd 
from selenium import webdriver
import os, time
from selenium.webdriver.common.keys import Keys

# %%

omgeving = input('Geef omgeving op (prod/test/opl) ')

# %%

if omgeving == 'prod':
    url = 'https://ssd.noordoostpoldernet.nl:442'
elif omgeving == 'test':
    url = 'https://ssdtest.noordoostpoldernet.nl:444'
elif omgeving == 'opl':
    url = 'https://ssdopl.noordoostpoldernet.nl:442'
else:
    raise Exception('Foutieve omgeving! Script stopt.')

# %%
streepje = '-'*20
print(streepje)
print('Het excelbestand mag van alles bevatten aan kolommen, maar er moet in ieder geval een kolom in zitten met de naam \'Aanvraagnummer\'. ')
print(streepje)
print()
df = pd.read_excel(input('Geef de naam van het in te lezen excelbestand (inclusief .xlsx)'))




# %%
options = webdriver.ChromeOptions()
options.add_argument("--log-level=3")
driver = webdriver.Chrome('D:\\python\\chromedriver\\chromedriver.exe',chrome_options=options)

driver.get(url)

# %%

input('Log in bij de Suites en druk daarna in dit scherm op Enter.')

# %%

driver.find_element_by_id('menuSearch').send_keys('intrekken werkproces')
driver.find_element_by_xpath('/html/body/ul/li[2]/a').click()

#292150
# %%

nietIngetrokken = 0
for werkproces in df[['Aanvraagnummer']].to_records():

    try:
        if driver.find_elements_by_id('M_C_SZAANVR__DD_AFDOENING_FASE') != []:
            driver.find_element_by_id('M_btnCancel').click()
            time.sleep(5)
    
    except:
        print(str(werkproces[1])+ ' is niet ingetrokken!')
    for i in range(1,8):
        driver.find_element_by_id('M_C_SZAANVR__AANVRAAGNR').send_keys(Keys.DELETE)
    time.sleep(2)
    try:
        driver.find_element_by_id('M_C_SZAANVR__AANVRAAGNR').send_keys(str(werkproces[1]))
        driver.find_element_by_id('M_C_SearchToolbar_btnSearch_defbut').click()
        try: 
            time.sleep(1)
            obj = driver.switch_to_alert()
            obj.accept()
        except:
            try:
                obj = driver.switch_to_alert()
                obj.accept()
            except:
                pass
        driver.find_element_by_id('M_btnSaveClose').click()
    except:
        print(str(werkproces[1])+ ' is niet ingetrokken!')
        nietIngetrokken += 1
    time.sleep(2)
   
# %%

driver.close()


print('Het script is klaar, er zijn ' + str(nietIngetrokken) + ' processen NIET ingetrokken!')
exit()

