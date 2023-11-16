# -*- coding: utf-8 -*-
"""
Created on Mon Feb 20 08:13:11 2022

@author: Januario Cipriano
"""

import re
import json
import time
import shutil
import logging
import threading
import numpy as np
import pandas as pd
from glob import glob
from pathlib import Path
from datetime import datetime
from selenium import webdriver
from multiprocessing import Process
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

from instantiate_driver import initiate_driver
from data_structure import ADT, NodeStructure
from classify_proposal import ClassifyProposal
from data_structure import ClientData, GetProcessNumber

client_data = ClientData()
process_numbers = GetProcessNumber()

adt = ADT()

"a202556"

"Business Banking"

def start_driver():
    URL = 'http://10.245.87.77/itcore.sites/ITCore.Ui.Web/ITCore.aspx'

    driver = initiate_driver()

    driver.get(URL)

    WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.ID, "MainIframe")))
    driver.switch_to.frame(driver.find_element(By.ID, "MainIframe"))

    dept_processes = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_fwcInit_BtnDepartmentWorklist"]')
    dept_processes.click()
    time.sleep(10)
    return driver

print(datetime.now())
driver = start_driver()

"a202556"

def fetch_page_links():
    link_ids = []
    for link_id in range(0, 30):
        try:
            link_path = f'//*[@id="ContentPlaceHolder1_flwForm_fwcInit_grvPager_repPager_pb1_{link_id}"]'
            link = driver.find_element(By.XPATH, link_path)
            process_numbers.add_link(link_path)
            # print(f'ADDING LINK OF PAGE {link_id}')
        except Exception as e:
            pass
            # print(e)
        time.sleep(0.1)

"a202556"

def fetch_process_numbers():
    print('LINKS')
    # print(process_numbers.link_ids)
    for link_id in process_numbers.link_ids:
        link = driver.find_element(By.XPATH, link_id)
        link.click()
        print(f'CLICKING ON PAGE {link_id}')
        time.sleep(8)
        
        for process_number in range(30):
            try:
                process_numbers_path = f'//*[@id="ContentPlaceHolder1_flwForm_fwcInit_gridAbertConta_lb_nProcesso_{process_number}"]'
                process_number_ele = driver.find_element(By.XPATH, process_numbers_path)
                client_name = driver.find_element(By.XPATH, f'//*[@id="ContentPlaceHolder1_flwForm_fwcInit_gridAbertConta_lb_nome_{process_number}"]')
                tipo = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_fwcInit_gridAbertConta_lb_tipoConta_6"]')
                segmento = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_fwcInit_gridAbertConta_lb_Segmento_0"]')
                process_numbers.add_process(
                    process_number_ele.text, client_name.text, tipo.text, segmento.text)
                print((process_number_ele.text + ': ' + client_name.text))
            except Exception as e:
                pass
                # print(e)
            time.sleep(0.2)
        time.sleep(3)
    print('PROCESSES')
    # print(process_numbers.process_numbers)


"a202556"

def check_if_public_client(words):
    words = words.lower()
    is_public_client = False
    if ('secto publico' in words or 'sector publico' in words or 
        'retencao na fonte' in words):
        is_public_client = True
    return is_public_client

"848276796"
"857995900"
"a202556"

def refresh_and_search(process_number):
    menu = driver.find_element(By.XPATH, '//*[@id="contentContainer_TMenus_rpCustMenu_LkCustMenu_4"]/div')
    hidden_submenu = driver.find_element(By.XPATH, '//*[@id="contentContainer_TMenus_rpCustMenu_rpSubCustMenu_4_rpSubCustMenuChilds_0_LkSubMenuChild_0"]')
      
    actions = ActionChains(driver)
    actions.move_to_element(menu)
    actions.click(hidden_submenu)
    actions.perform()
    time.sleep(0.1)

    print('WAITING FOR VISIBILITY OF MainIframe')
    WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "MainIframe")))
    driver.switch_to.frame(driver.find_element(By.ID, "MainIframe"))
    print('MainIframe ALREADY VISIBLE')

    print('WAITING FOR VISIBILITY OF search_process')
    WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_ntbcProcessNumber_txField"]')))
    search_process_field_XPATH = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_ntbcProcessNumber_txField"]')
    search_process_field_XPATH.send_keys(str(process_number))
    time.sleep(0.5)
    logging.warning('Successfully inserted PROCESS NUMBER into input field')
    
    print('ATTEMPTING TO CLICK ON SEARCH BUTTON')
    search_button_XPATH = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_fcontrols_btnSearch"]/span[2]')
    search_button_XPATH.click()
    print('SEARCH BUTTON CLICKED')
    time.sleep(8)

    print('WAITING FOR VISIBILITY OF CLIENT TD')
    WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_egvSearchResult"]/tbody/tr[2]')))
    client_td = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_egvSearchResult"]/tbody/tr[2]')
    client_td.click()
    print('CLIENT TD CLICKED')
    time.sleep(3)


"857995900"
"a202556"


def get_row():
    driver.refresh()
    time.sleep(6)

    process_numbers_ = process_numbers.process_numbers
    print(process_numbers_)

    for process_number, client_name, tipo, segmento in process_numbers_:
        client_data.process_number = process_number
        client_data.name = client_name
        client_data.tipo = tipo
        client_data.segmento = segmento
        print(f'PROCESS NUMBER = {process_number}')
        # refresh_and_search(process_number)

        menu = driver.find_element(By.XPATH, '//*[@id="contentContainer_TMenus_rpCustMenu_LkCustMenu_4"]/div')
        hidden_submenu = driver.find_element(By.XPATH, '//*[@id="contentContainer_TMenus_rpCustMenu_rpSubCustMenu_4_rpSubCustMenuChilds_0_LkSubMenuChild_0"]')
          
        actions = ActionChains(driver)
        actions.move_to_element(menu)
        actions.click(hidden_submenu)
        actions.perform()
        time.sleep(0.1)

        print('WAITING FOR VISIBILITY OF MainIframe')
        WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "MainIframe")))
        driver.switch_to.frame(driver.find_element(By.ID, "MainIframe"))
        print('MainIframe ALREADY VISIBLE')

        print('WAITING FOR VISIBILITY OF search_process')
        WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_ntbcProcessNumber_txField"]')))
        search_process_field_XPATH = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_ntbcProcessNumber_txField"]')
        search_process_field_XPATH.send_keys(str(process_number))
        time.sleep(0.5)
        logging.warning('Successfully inserted PROCESS NUMBER into input field')
        
        print('ATTEMPTING TO CLICK ON SEARCH BUTTON')
        search_button_XPATH = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_fcontrols_btnSearch"]/span[2]')
        search_button_XPATH.click()
        print('SEARCH BUTTON CLICKED')
        # time.sleep(6.5)

        print('WAITING FOR VISIBILITY OF CLIENT TD')
        td_path = '//*[@id="ContentPlaceHolder1_flwForm_egvSearchResult"]/tbody/tr[2]'
        WebDriverWait(driver, 45).until(EC.visibility_of_element_located((By.XPATH, td_path)))
        client_td = driver.find_element(By.XPATH, td_path)
        client_td.click()
        print('CLIENT TD CLICKED')
        # time.sleep(8)

        try:
            WebDriverWait(driver, 45).until(EC.visibility_of_element_located((By.XPATH, '/html/body/form/div[3]/section/span[2]/div[1]/div[1]/div/div[3]/ul/li[2]/a')))

            print('ATTEMPTING TO CLICK ON DADOS FINANCEIROS')
            financial_data = driver.find_element(By.XPATH, '/html/body/form/div[3]/section/span[2]/div[1]/div[1]/div/div[3]/ul/li[2]/a')
            financial_data.click()
            print('DADOS FINANCEIROS CLICKED')
            time.sleep(0.5)

            requested_amount = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ff1_fc1_tabs_fic1b_txt_GoodValue"]')
            requested_amount = re.findall(r'\d+', requested_amount.text)
            last = requested_amount[-1]
            requested_amount = ''.join(requested_amount[:-1]) + '.'+ last
            print(requested_amount)
            requested_amount = float(requested_amount)
            client_data.requested_amount = requested_amount
            print(f'Valor Requisitado: {requested_amount}')

            interest_rate = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ff1_fc1_tabs_fic1b_DgTaxasCredito_ctl02_txtDGtaxaspread"]')
            interest_rate = float(interest_rate.get_attribute('value').replace(',', '.'))
            client_data.interest_rate = interest_rate
            print(f'Interest Rate: {interest_rate}')

            conditions = driver.find_element(By.XPATH, '/html/body/form/div[3]/section/span[2]/div[1]/div[1]/div/div[3]/ul/li[4]/a')
            conditions.click()
            time.sleep(0.3)

            # business_conditions = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_ff1_fc1_tabs_fic1Conditions_txtBusinessConditions')
            business_conditions = driver.find_element(By.XPATH, '//span[@id="ctl00_ContentPlaceHolder1_ff1_fc1_tabs_fic1Conditions_txtBusinessConditions"]')
            business_conditions = business_conditions.get_attribute("textContent")
            client_data.business_conditions = business_conditions
            time.sleep(0.2)

            proponents = driver.find_element(By.XPATH, '/html/body/form/div[3]/section/span[2]/div[1]/div[2]/ul[2]/li[4]')
            proponents.click()
            time.sleep(0.2)

            # client_name = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ff1_fc1_finTit_dgInterveners_ctl02_lblName"]')
            # client_data.name = client_name.text

            verificar_proponents = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ff1_fc1_finTit_dgInterveners_ctl02_btnDetailIntervener"]')
            verificar_proponents.click()
            time.sleep(0.2)

            driver.switch_to.parent_frame()
            time.sleep(0.2)

            try:
                WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, 'PopupMainIframeShowAddIntervener_Popup')))
                driver.switch_to.frame(driver.find_element(By.ID, 'PopupMainIframeShowAddIntervener_Popup'))

                profissao_entidade_patronal = driver.find_element(By.XPATH, '//*[@id="ui-id-5"]/span')
                profissao_entidade_patronal.click()
                time.sleep(0.2)
                
                entidade_patronal = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_PopupContainer_detalhePropop_tabs_fic4c_txtProfessionalEmployer"]')
                entidade_patronal = entidade_patronal.text
                client_data.entidade_patronal = entidade_patronal

                entidade_empregadora = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_PopupContainer_detalhePropop_tabs_fic4c_ddlEntidadeEmpregadora"]')
                entidade_empregadora = entidade_empregadora.text
                client_data.entidade_empregadora = entidade_empregadora

                # if entidade_patronal == '' or entidade_patronal == 'Outra':
                #     client_data.entidade_patronal = entidade_empregadora
                    
                # print(json.dumps(client_data.values, default=str, indent=4))
                # client_data.add()
                # print(f'BUSINESS CONDITIONS: {business_conditions}')
                print(f'check_if_public_client({business_conditions}) = {check_if_public_client(business_conditions)}')
                print(f'interest_rate = {interest_rate}')

                if requested_amount > 500000:
                    print('Amount greater than 500k')
                else:
                    print('Amount less than or equal to 500k')
                    if interest_rate == 5 and check_if_public_client(business_conditions):
                        print('Interest rate equal to 5%. Funcionario Publico')
                    else:
                    # elif interest_rate != 5 and not check_if_public_client(business_conditions):
                        print('Libertar processo')
                        client_data.libertar = True
                print(json.dumps(client_data.values, default=str, indent=4))
                    # else:
                    #     print('None of the above')
                    # else:
                    #     client_data.libertar = True
                    #     print('Libertar o processo')

                if client_data.libertar is True:
                    print('LIBERTAR PROCESSO')
                    # driver.refresh()
                    # time.sleep(6)

                    # refresh_and_search(process_number)

                    # workflow_path = '//*[@id="st_16"]/span[1]'
                    # WebDriverWait(driver, 40).until(EC.visibility_of_element_located((By.XPATH, workflow_path)))
                    # workflow = driver.find_element(By.XPATH, workflow_path)
                    # print(workflow.text)
                    # print('CLICKING WORKFLOW - SECOND TIME')
                    # workflow.click()
                    # print('WORKFLOW CLICKED - SECOND TIME')
                    # time.sleep(7)

                    # driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.HOME)

                    # proxima_acao = driver.find_element(By.XPATH, '/html/body/form/div[3]/section/span[2]/div[1]/div[1]/div/div[3]/ul/li[2]/a')
                    # proxima_acao.click()
                    # print('CLICKED PROXIMA ACCAO - SECOND TIME')
                    # time.sleep(2)

                    # libertar_processo = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ff1_fc1_tabs_procValidationControl_fic_ficValidation_rptRadios_ctl02_rbAction_cbField"]')
                    # libertar_processo.click()
                    # time.sleep(2)

                    # driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.PAGE_DOWN)
                    # time.sleep(2.5)
                else:
                    print('NAO LIBERTAR PROCESSO')

            except Exception as e:
                print(e)
        except Exception as e:
            print(e)

        client_data.add()
        driver.refresh()
        # client_data.clear_values()
        time.sleep(6)
        client_data.save_to_excel()
        client_data.clear_values()

    print(pd.DataFrame(client_data.data))
    client_data.send_email()
    # client_data.save_to_excel()

"a202556"


# 863885151
# 86/87-5610610


"a202556"

def refresh_to_initial_page():
    print('REFRESHING PAGE...')
    driver.refresh()
    print('PAGE SUCCESSFULLY REFRESHED')
    time.sleep(4)

"a202556"


def fetch_table(page_id):
    logging.info(f"Fetching data at page {page_id}")
    logging.warning(f'Fetching page {adt.PAGE_COUNT}')
    adt.PAGE_COUNT += 1
    for path_id in range(0, 15):
        try:
            start = datetime.now().strftime('%d-%b-%Y %H:%M:%S')

            get_row(path_id)
            time.sleep(1)
            end = datetime.now().strftime('%d-%b-%Y %H:%M:%S')
            logging.warning(f'Fetching row {path_id+1}. \tStarted at {start}. \tEnded at: {end}')
            # driver.save_screenshot('Page1.png')
        except Exception as exc:
            print(exc)
            logging.warning(exc)
    logging.warning('\n')
    time.sleep(3)
    print('\n')

def fetch_remaining_pages_2(page_id):
    link = driver.find_element(By.XPATH, f'//*[@id="ContentPlaceHolder1_flwForm_fwcInit_grvPager_repPager_pb1_{page_id}"]')
    link.click()


"a202556"

        
def fetch_remaining_pages():
    link_ids = []
    for link_id in range(1, 30):
        try:
            link = driver.find_element(By.XPATH, f'//*[@id="ContentPlaceHolder1_flwForm_fwcInit_grvPager_repPager_pb1_{link_id}"]')
            link_ids.append(link_id)
        except Exception:
            pass

    for link_id in link_ids:
        print(f'\nFETCHING PAGE = {link_id}')
        try:
            link = driver.find_element(By.XPATH, f'//*[@id="ContentPlaceHolder1_flwForm_fwcInit_grvPager_repPager_pb1_{link_id}"]')
            link.click()
            time.sleep(8)
            fetch_table(link_id)
            # time.sleep(6)
            save_to_excel()
            time.sleep(2)
        except Exception as exc:
            print(exc)
            logging.warning(exc)
    

"a202556"


def save_to_excel():
    logging.info('Saving data to excel file')
    df = pd.DataFrame(adt.data)
    df.columns = 'SLA, Nº, Tipo, Nome, Segmento, Estado, Balcão de Criação, Colaborador, R, Dt. Criação, DownloadedAt, Valor Requisitado, Area de Verificacao, Entidade Patronal, IsUpdated, IsPropostaActualizada'.split(', ')
    now = datetime.now()
    if now.hour < 12:
        now = now.date().strftime('%d.%m.%Y')
        filename = f'Fluxo - {now} (Manha).xlsx'
    else:
        now = now.date().strftime('%d.%m.%Y')
        filename = f'Fluxo - {now} (Tarde).xlsx'

    df.drop_duplicates('Nº', inplace=True)

    df.to_excel(filename, index=False)
    time.sleep(2)

"a202556"


def main():
    time.sleep(0.5)
    fetch_page_links()
    fetch_process_numbers()
    time.sleep(1.5)

    get_row()


"a202556"

def fetch_table_page(page_id):
    for td_id in range(0, 15):
        try:
            # driver = start_driver()
            link = driver.find_element(By.XPATH, f'//*[@id="ContentPlaceHolder1_flwForm_fwcInit_grvPager_repPager_pb1_{page_id}"]')
            print(f'LINK ID: {link.text}')
            link.click()
            time.sleep(7)

            get_row(td_id)
        except Exception as e:
            pass


def run_all():
    for i in range(6):
        t = threading.Thread(target=fetch_table_page, args =[i])
        t.start()


"a202556"

if __name__=='__main__':
    start_time = time.perf_counter()
    print(datetime.now())
    main()

    # from search_process import search_client_process 

    # print('SWITCHING TO PARENT NODE================================')
    # driver.switch_to.parent_frame()
    # print('STARGING TO SEARCH PROPOSALS============================')
    # for i in range(4):
    #     search_client_process(driver)
    #     time.sleep(2)

    # from split_proposal_types import split_proposal
    # split_proposal()

    # elapsed_time = time.perf_counter() - start_time
    # print(f'ELAPSED TIME = {elapsed_time}')
    # print(datetime.now())

"a202556"































