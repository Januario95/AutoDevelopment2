import os
import json
import string
import shutil
import numpy as np
import pandas as pd
from glob import glob
from enum import Enum
import win32com.client
from datetime import datetime

share_folder = '//10.245.10.81/Direccao Financeira/EDO/3. INTELLIGENT AUTOMATION/Credit Card/'


class GetProcessNumber:
    def __init__(self):
        self.process_numbers = []
        self.link_ids = []

    def add_link(self, link):
        self.link_ids.append(link)

    def add_process(self, process_number, client_name, tipo, segmento):
        # tipo = tipo.translate(str.maketrans('', '', string.punctuation))
        self.process_numbers.append((process_number, client_name, tipo, segmento))

class ClientData:
    def __init__(self):
        self.name = None
        self.process_number = None
        self.tipo = None
        self.segmento = None
        self.requested_amount = None
        self.interest_rate = None
        self.business_conditions = None
        self.entidade_patronal = None
        self.entidade_empregadora = None
        self.libertar = False
        self.data = []

    def add(self):
        self.data.append(self.values)

    def clear_values(self):
        self.name = None
        self.process_number = None
        self.tipo = None
        self.segmento = None
        self.requested_amount = None
        self.interest_rate = None
        self.business_conditions = None
        self.entidade_patronal = None
        self.entidade_empregadora = None
        self.libertar = False

    @property
    def values(self):
        return {
            'name': self.name,
            'process_number': self.process_number,
            'tipo': self.tipo,
            'segmento': self.segmento,
            'requested_amount': self.requested_amount,
            'interest_rate': self.interest_rate,
            'business_conditions': self.business_conditions,
            'entidade_patronal': self.entidade_patronal,
            'entidade_empregadora': self.entidade_empregadora,
            'libertar': self.libertar
        }

    def send_email(self):
        print('ATTEMPTING TO SEND EMAIL')
        attachment = os.getcwd() + '\\Extracao do Fluxo.xlsx'
        engine = win32com.client.Dispatch('outlook.application')
        email = engine.CreateItem(0)
        # email.Recipients.Add('conceicao.monteiro@standardbank.co.mz')
        email.Recipients.Add('januario.cipriano@standardbank.co.mz')
        email.Subject = 'Extracao do Fluxo'
        email.HtmlBody = f'<h4 style="text-align:center;font-style:italic;">Extracao do Fluxo - Collateral</h4>'
        email.attachments.Add(attachment)
        email.send
        print('EMAIL SENT SUCCESSFULLY')

    def save_to_excel(self):  
        # data = {key: [val] for key, val in self.values.items()}
        df = pd.DataFrame(self.data)
        df.columns = ['Nome', 'Numero do Processo', 'Tipo', 'Segmento', 'Montante Solicitado',
                      'Taxa de Juro (%)', 'Condicoes de Negocio', 'Entidade Empregadora',
                      'Entidade Empregadora', 'Libertar']
        df.to_excel(f'{os.getcwd()}\\Extracao do Fluxo.xlsx', index=False)


# client_data = ClientData()
# client_data.send_email()

class NodeStructure:
    COUNT = 0
    def __init__(self, timer, nProcesso, tipoConta, nome, Segmento,
                 estado, branch, user, R, CreationDate, ModificadoEm,
                 ValorRequisitado, Colaborador2, EntidadePatronal, IsUpdated,
                 IsPropostaActualizada):
        self.timer = timer
        self.nProcesso = nProcesso
        self.tipoConta = tipoConta
        self.nome = nome
        self.Segmento = Segmento
        self.estado = estado
        self.branch = branch
        self.user = user
        self.R = R
        self.CreationDate = CreationDate
        self.ModificadoEm = ModificadoEm
        self.ValorRequisitado = ValorRequisitado
        self.Colaborador2 = Colaborador2
        self.EntidadePatronal = EntidadePatronal
        self.IsUpdated = IsUpdated
        self.IsPropostaActualizada = IsPropostaActualizada
        NodeStructure.COUNT += 1
        
    def __str__(self):
        return f'Node-{self.COUNT}'
        
    @property
    def values(self):
        return {
            'sla': self.timer,
            'process_number': self.nProcesso,
            'type': self.tipoConta, 
            'name': self.nome,
            'segment': self.Segmento,
            'state': self.estado,
            'branch': self.branch,
            'user': self.user,
            'R': self.R,
            'creation_date': self.CreationDate, 
            'Modificado-Em': self.ModificadoEm,
            'ValorRequisitado': self.ValorRequisitado,
            'Colaborador2': self.Colaborador2,
            'EntidadePatronal':  self.EntidadePatronal,
            'IsUpdated': self.IsUpdated,
            'IsPropostaActualizada': self.IsPropostaActualizada
        }
        

class ADT:
    def __init__(self, nodes=[]):
        self.__data = []
        self.PAGE_COUNT = 0
        self.IS_PROPOSTA_ATUALIZADA = False
        for node in nodes:
            self.__raise_on_error(node)
            self.__data.append(node)
            
    def __str__(self):
        return f'<ADT ({len(self.__data)} Nodes)>'
    
    def __len__(self):
        return len(self.__data)
        
    def push(self, nodes):
        if isinstance(nodes, list):
            for node in nodes:
                self.__raise_on_error(node)
                self.__data.append(node)
        else:
            self.__raise_on_error(nodes)
            self.__data.append(nodes)
            
    def __raise_on_error(self, node):
        if not isinstance(node, NodeStructure):
            raise TypeError('Node must be an instance of NodeStructure')
        
    @property
    def data(self):
        return [node.values for node in self.__data]

    




    



















