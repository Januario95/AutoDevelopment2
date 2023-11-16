# -*- coding: utf-8 -*-
"""
Created on Fri Oct 14 11:57:12 2022

@author: Januario Cipriano
"""

import re
import os
import json
import time
import PyPDF2
import shutil
import logging
import pyautogui
import pdfplumber
import numpy as np
import pandas as pd
from glob import glob
from PIL import Image
from pathlib import Path
import aspose.words as aw
from datetime import datetime
from selenium import webdriver
from multiprocessing import Process
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

from instantiate_driver import initiate_driver
from data_structure import ClientData, GetProcessNumber

client_data = ClientData()
project_path = os.getcwd()
process_numbers = GetProcessNumber()

"a202556"


URL = 'http://10.245.87.77/itcore.sites/ITCore.Ui.Web/ITCore.aspx'
share_folder = '//10.245.10.81/Direccao Financeira/EDO/3. INTELLIGENT AUTOMATION/Credit Card'
downloads_folder = '/'.join(os.getcwd().split('\\')[:3]) + '/Downloads'

print('=' * 50)
print(datetime.now())
print('=' * 50)
driver = initiate_driver()

driver.get(URL)
time.sleep(4)

"a202556"

WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.ID, "MainIframe")))
driver.switch_to.frame(driver.find_element(By.ID, "MainIframe"))

dept_processes = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_fwcInit_BtnDepartmentWorklist"]')
dept_processes.click()
time.sleep(10)


"a202556"


def extract_data():
    os.chdir(downloads_folder)
    list_of_files = glob('*.pdf')
    latest_file = max(list_of_files, key=os.path.getctime)
    filename = f'{downloads_folder}/{latest_file}'
    reader = PyPDF2.PdfReader(filename)
    text = reader.pages[0].extract_text()
    text = text.split('\n')[12:]
    os.chdir(project_path)

    df = pd.read_csv('Branches.csv')

    data = {}
    for index in range(0, len(text), 2):
        try:
            # print(f'{index} = {text[index]}: {text[index+1]}')

            if 'Data do Limite' in text[index+1]:
                nr_de_conta = text[index]
                client_data.Conta_de_Bebito = nr_de_conta

            if 'Limite de crédito por extenso' in text[index+1]:
                limit_de_credito = re.findall(r'\d+', text[index])
                limit_de_credito = ''.join(limit_de_credito[:-1]) + '.'+ limit_de_credito[-1]
                client_data.Limite = limit_de_credito

            if 'A ser preenchido pelo banco' in text[index+1]:
                balcao_de_entrega = text[index]
                balcao_de_entrega = balcao_de_entrega.split(' ')[-1].upper()
                balcao_de_entrega = df[df['Branch'].str.contains(balcao_de_entrega)]['Short Number'].tolist()[0]
                client_data.Balcao_De_Entrega = balcao_de_entrega

            if 'Conta número' in text[index+1]:
                modalidade_de_pag = text[index]
                client_data.Modalidade_De_Pag = modalidade_de_pag

            if 'Conta a debitar no dia' in text[index]:
                data_de_debito = text[index+1]
                client_data.Data_De_Debito_Directo = data_de_debito
            else:
                pass
        except IndexError:
            pass
    os.chdir(project_path)
    print(json.dumps(
        client_data.values,
        indent=4))

"a202556"

def create_client_folder():
    folder = share_folder + f'/documents/{client_data.Nome}/'
    if not os.path.exists(folder):
        os.mkdir(folder)


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
    print(process_numbers.link_ids)
    for link_id in process_numbers.link_ids:
        link = driver.find_element(By.XPATH, link_id)
        link.click()
        print(f'CLICKING ON PAGE {link_id}')
        
        for process_number in range(30):
            try:
                print(f'ATTEMPTING TO FETCH DATA ON PAGE {link_id}')
                process_numbers_path = f'//*[@id="ContentPlaceHolder1_flwForm_fwcInit_gridAbertConta_lb_nProcesso_{process_number}"]'
                element = driver.find_element(By.XPATH, process_numbers_path)
                process_numbers.add_process(element.text)
                # print(element.text)
            except Exception as e:
                pass
            time.sleep(0.2)
        time.sleep(3)
    print('PROCESSES')
    print(process_numbers.process_numbers)

"a202556"

"Sharmila da Conceicao Abencoado Macuacua"

def get_row():
    driver.refresh()
    time.sleep(5)

    process_numbers_ = process_numbers.process_numbers
    print(process_numbers_)

    for process_number in process_numbers_:
        print(f'PROCESS NUMBER = {process_number}')
        menu = driver.find_element(By.XPATH, '//*[@id="contentContainer_TMenus_rpCustMenu_LkCustMenu_3"]/div')
        hidden_submenu = driver.find_element(By.XPATH, '//*[@id="contentContainer_TMenus_rpCustMenu_rpSubCustMenu_3_rpSubCustMenuChilds_0_LkSubMenuChild_0"]')
              
        actions = ActionChains(driver)
        actions.move_to_element(menu)
        actions.click(hidden_submenu)
        actions.perform()
        time.sleep(2)
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.HOME)
        logging.warning('Successfully clicked on search href.')

        WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "MainIframe")))
        driver.switch_to.frame(driver.find_element(By.ID, "MainIframe"))

        WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_ntbcProcessNumber_txField"]')))
        search_process_field_XPATH = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_ntbcProcessNumber_txField"]')
        search_process_field_XPATH.send_keys(str(process_number))
        time.sleep(0.5)
        logging.warning('Successfully inserted PROCESS NUMBER into input field')
        
        search_button_XPATH = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_fcontrols_btnSearch"]/span[2]')
        search_button_XPATH.click()
        time.sleep(6)

        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_egvSearchResult"]/tbody/tr[2]')))
        client_td = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_egvSearchResult"]/tbody/tr[2]')
        client_td.click()
        time.sleep(3)

        first_window = driver.window_handles[0]

        emitir_contrato = None
        try:
            emitir_contrato = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ff1_fc1_tabs_fcc100_btnContrato"]/span[2]')
            emitir_contrato.click()
            print('CLICKED ON emitir_contrato FIRST-ATTEMPT')
            time.sleep(2.5)
        except Exception as e:
            pass

        if emitir_contrato is None:
            try:
                emitir_contrato = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ff1_fc1_tabs_fcc100_btnEmitirProposta"]')
                emitir_contrato.click()
                print('CLICKED ON emitir_contrato SECOND-ATTEMPT')
                time.sleep(2.5)
            except Exception as e:
                pass

        print(f'NUMBER OF OPENED WINDOWS: {len(driver.window_handles)}')
        try:
            second_window = driver.window_handles[1]
        except Exception as e:
            print(e)
            time.sleep(6)
            second_window = driver.window_handles[1]

        driver.switch_to.window(second_window)
        time.sleep(1)
        driver.maximize_window()
        time.sleep(1)

        pyautogui.hotkey('ctrl', 's')
        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(1)

        driver.switch_to.window(first_window)
        time.sleep(1)

        WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.ID, "MainIframe")))
        driver.switch_to.frame(driver.find_element(By.ID, "MainIframe"))

        path = '/html/body/form/div[3]/section/span[2]/div[1]/div[2]/ul[2]/li[4]/span[1]'
        wait = WebDriverWait(driver, 20)
        wait.until(EC.visibility_of_element_located((By.XPATH, path)))

        proponents = driver.find_element(By.XPATH, path)
        proponents.click()
        time.sleep(1.5)

        client_name = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ff1_fc1_finTit_dgInterveners_ctl02_lblName"]')
        client_data.Nome = client_name.text
        time.sleep(0.5)

        filename = f'{downloads_folder}/ver_documento.pdf'
        # print(filename)
        if os.path.exists(filename):
            extract_data()
            time.sleep(0.8)
            create_client_folder()
            time.sleep(0.3)
            shutil.move(filename, share_folder + f'/documents/{client_data.Nome}')

        proponents_box = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ff1_fc1_finTit_dgInterveners_ctl02_btnDetailIntervener"]')
        proponents_box.click()
        time.sleep(1)

        driver.switch_to.parent_frame()
        time.sleep(0.2)

        WebDriverWait(driver, 25).until(EC.visibility_of_element_located((By.ID, 'PopupMainIframeShowAddIntervener_Popup')))
        driver.switch_to.frame(driver.find_element(By.ID, 'PopupMainIframeShowAddIntervener_Popup'))
        time.sleep(1)

        client_number = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_PopupContainer_detalhePropop_tabs_fic1a_lblCIF"]')
        client_data.Nr_Do_Cliente = client_number.text

        date_of_birth = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_PopupContainer_detalhePropop_tabs_fic1a_dtHolderBirthDate"]')
        client_data.Data_De_Nascimento = date_of_birth.text.replace('-', '/')

        marital_status = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_PopupContainer_detalhePropop_tabs_fic1a_ddlHolderMaritalStatus"]')
        client_data.Estdo_Civil = marital_status.text

        nationality = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_PopupContainer_detalhePropop_tabs_fic1a_rbHolderIsMozambicanY"]')
        if nationality.text == 'Sim':
            client_data.Nacionalidade = 'Mocambique'

        sexo = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_PopupContainer_detalhePropop_tabs_fic1a_rbHolderGenderM"]')
        client_data.Sexo = sexo.text

        identification_nr_tab = driver.find_element(By.XPATH, '/html/body/form/div[3]/section/span[2]/div[2]/table/tbody/tr/td/div/div[1]/ul/li[2]/a')
        identification_nr_tab.click()
        time.sleep(0.3)

        nr_do_BI = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_PopupContainer_detalhePropop_tabs_fic1b_txtHolderDocumentNumber"]')
        client_data.Nr_Do_BI = nr_do_BI.text

        morada_tab = driver.find_element(By.XPATH, '/html/body/form/div[3]/section/span[2]/div[2]/table/tbody/tr/td/div/div[1]/ul/li[4]/a')
        morada_tab.click()
        time.sleep(0.3)

        Endereco = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_PopupContainer_detalhePropop_tabs_fic3a_txtAddressNeighborhood"]')
        client_data.Endereco = Endereco.text[:30]

        province = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_PopupContainer_detalhePropop_tabs_fic3a_ddlAddressProvince"]')
        client_data.Localicade_e_Provincia = province.text

        profession_and_employer_tab = driver.find_element(By.XPATH, '/html/body/form/div[3]/section/span[2]/div[2]/table/tbody/tr/td/div/div[1]/ul/li[5]/a')
        profession_and_employer_tab.click()
        time.sleep(0.3)



        # driver.save_screenshot('PROFISSAO-ENT_PATRONAL.png')
        # time.sleep(2)
        # img = Image.open('PROFISSAO-ENT_PATRONAL.png')
        # box = (300, 0, 1600, 400)
        # img2 = img.crop(box)
        # os.remove('PROFISSAO-ENT_PATRONAL.png')
        # time.sleep(0.5)
        # img2.save('PROFISSAO-ENT_PATRONAL.png')

        segmento = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_PopupContainer_detalhePropop_tabs_fic4a_ddlTargetCentral"]')
        client_data.Segmento = segmento.text.split('-')[-1].upper()

        contact_tab = driver.find_element(By.XPATH, '/html/body/form/div[3]/section/span[2]/div[2]/table/tbody/tr/td/div/div[1]/ul/li[6]/a')
        contact_tab.click()
        time.sleep(0.3)

        # driver.save_screenshot('CONTACTOS.png')
        # img = Image.open('CONTACTOS.png')
        # box = (300, 0, 1600, 1000)
        # img2 = img.crop(box)
        # os.remove('CONTACTOS.png')
        # time.sleep(0.5)
        # img2.save('CONTACTOS.png')

        phone_number = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_PopupContainer_detalhePropop_tabs_fic5b_lbl_txtContactNatMobile"]')
        phone_number = phone_number.text
        if len(phone_number) == 9:
            phone_number = '+258' + phone_number
        client_data.Contactos = phone_number

        client_data.save_to_excel()
        print(client_data.values)
        time.sleep(0.5)

        driver.refresh()
        time.sleep(5)

        menu = driver.find_element(By.XPATH, '//*[@id="contentContainer_TMenus_rpCustMenu_LkCustMenu_3"]/div')
        hidden_submenu = driver.find_element(By.XPATH, '//*[@id="contentContainer_TMenus_rpCustMenu_rpSubCustMenu_3_rpSubCustMenuChilds_0_LkSubMenuChild_0"]')
              
        actions = ActionChains(driver)
        actions.move_to_element(menu)
        actions.click(hidden_submenu)
        actions.perform()
        time.sleep(1)
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.HOME)
        logging.warning('Successfully clicked on search href.')

        WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "MainIframe")))
        driver.switch_to.frame(driver.find_element(By.ID, "MainIframe"))

        WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_ntbcProcessNumber_txField"]')))
        search_process_field_XPATH = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_ntbcProcessNumber_txField"]')
        search_process_field_XPATH.send_keys(str(process_number))
        time.sleep(0.2)
        logging.warning('Successfully inserted PROCESS NUMBER into input field')
        
        search_button_XPATH = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_fcontrols_btnSearch"]/span[2]')
        search_button_XPATH.click()
        time.sleep(5)

        WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_egvSearchResult"]/tbody/tr[2]')))
        client_td = driver.find_element(By.XPATH, '//*[@id="ContentPlaceHolder1_flwForm_egvSearchResult"]/tbody/tr[2]')
        client_td.click()
        time.sleep(2.5)

        WebDriverWait(driver, 25).until(EC.visibility_of_element_located((By.XPATH, "/html/body/form/div[3]/section/span[2]/div[1]/div[2]/ul[2]/li[9]")))

        document_checklist = driver.find_element(By.XPATH, '/html/body/form/div[3]/section/span[2]/div[1]/div[2]/ul[2]/li[9]')
        document_checklist.click()

        time.sleep(1.5)

        BI_Comprovativo = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ff1_fc1_finTit_finTit2_rptValidationSummary_ctl01_rptValidationSummaryDetails_ctl00_lblDetailInfoLink"]')
        BI_Comprovativo.click()

        third_window = driver.window_handles[2]
        driver.switch_to.window(third_window)
        time.sleep(1)
        driver.maximize_window()
        time.sleep(1)

        names = client_data.Nome.title().split(' ')
        name = names[0] + ' ' + names[-1]
        filename = downloads_folder + '/' + name + '  BI.pdf'
        # list_of_files = glob('*.pdf')
        # latest_file = max(list_of_files, key=os.path.getctime)
        os.chdir(downloads_folder)
        if os.path.exists(filename):
            print(filename)
            print('=' * 50)
            print('File Exists')
            shutil.move(filename, share_folder + f'/documents/{client_data.Nome}')
            print('=' * 50)
        else:
            print(filename)
            print('File Does Not Exist')

        filename = downloads_folder + '/B.I..pdf'
        if os.path.exists(filename):
            print(filename)
            print('=' * 50)
            print('File Exists')
            shutil.move(filename, share_folder + f'/documents/{client_data.Nome}')
            print('=' * 50)
        else:
            print(filename)
            print('File Does Not Exist')
        time.sleep(1)
        os.chdir(project_path)

        driver.switch_to.window(first_window)
        time.sleep(1)

        WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.ID, "MainIframe")))
        driver.switch_to.frame(driver.find_element(By.ID, "MainIframe"))

        path = '//*[@id="ctl00_ContentPlaceHolder1_ff1_fc1_fcc_btnBackToProp"]'
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, path)))

        back_to_previous_page = driver.find_element(By.XPATH, path)
        back_to_previous_page.click()
        time.sleep(4)

        workflow_path = '//*[@id="st_16"]/span[1]'
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.XPATH, workflow_path)))
        workflow = driver.find_element(By.XPATH, workflow_path)
        print(workflow.text)
        print('CLICKING WORKFLOW')
        workflow.click()
        print('WORKFLOW CLICKED')
        time.sleep(5)

        print('ATTEMPTING TO SCROLL TO TOP OF THE PAGE')
        driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL + Keys.HOME)
        print('SUCCESSFULLY SCROLLED TO THE TOP')

        print('WAITING FOR NEXT ACTION TO BE VISIBLE')
        next_action_path = '/html/body/form/div[3]/section/span[2]/div[1]/div[1]/div/div[3]/ul/li[2]/a'
        WebDriverWait(driver, 8).until(EC.visibility_of_element_located((By.XPATH, next_action_path)))
        print('NEXT ACTION ALREADY VISIBLE IN THE PAGE')
        next_action = driver.find_element(By.XPATH, next_action_path)
        print('NEXT ACTION CAPTURED BY THE DRIVER')
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, next_action_path)))
        print('AWAITING FOR ELEMENT TO BE CLICKABLE')
        time.sleep(1)


        try:
            print('ATTEMPTING TO CLICK ON NEXT ACTION')
            next_action.click()
            time.sleep(1)

            view_more = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ff1_fc1_tabs_procValidationControl_fic_ficComments"]/div[2]/a')
            view_more.click()
            time.sleep(0.5)

            gestora = None
            data_de_criacao = None
            for path_id in range(1, 20):
                try:
                    path = f'//*[@id="ctl00_ContentPlaceHolder1_ff1_fc1_tabs_procValidationControl_fic_ficComments_rptComments_ctl0{path_id}_lblColaborador"]'
                    elem = driver.find_element(By.XPATH, path)
                    gestora = elem.text
                except Exception as e:
                    pass
                    
                try:
                    data_de_criacao = driver.find_element(By.XPATH, '//*[@id="ctl00_ContentPlaceHolder1_ff1_fc1_tabs_procValidationControl_fic_ficComments_rptComments_ctl05_lblData"]')
                    data_de_criacao = data_de_criacao.text
                    print('=' * 50)
                    print(f'GESTOR = {gestora}')
                    print('=' * 50)
                except Exception as e:
                    pass
            if gestora is not None:
                client_data.Gestor_A = gestora
            if data_de_criacao is not None:
                client_data.Data_Da_Criacao_No_Wf = data_de_criacao
                client_data.save_to_excel()
                print(client_data.values)
            time.sleep(1)

        except Exception as e:
            print('ERROR ATTEMPTING TO CLICK ON NEXT ACTION')
            print(e)
        time.sleep(1)

        driver.refresh()
        client_data.clear_values()
        time.sleep(5)


"a202556"


def main():
    time.sleep(0.5)
    fetch_page_links()
    fetch_process_numbers()
    time.sleep(1.5)

    get_row()

    # doc = aw.Document()
    # create a document builder object
    # builder = aw.DocumentBuilder(doc)

    # builder.write("\n\n")
    # builder.insert_image("PROFISSAO-ENT_PATRONAL.png")
    # builder.write("\n\n")
    # builder.insert_image("CONTACTOS.png")
    # doc.save('PROFISSAO-ENT_PATRONAL.docx')


    # with pdfplumber.open(filename) as pdf:
    #     first_page = pdf.pages[0]
    #     print(first_page.extract_text())

"a202556"

if __name__=='__main__':
    start_time = time.perf_counter()
    main()

    print('=' * 50)
    print(datetime.now())
    print('=' * 50)
    elapsed_time = time.perf_counter() - start_time
    print(f'ELAPSED TIME = {elapsed_time}')

"a202556"































