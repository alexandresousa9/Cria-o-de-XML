#!/usr/bin/python
# -*- coding: UTF-8 -*-

from xml.dom.minidom import parse, parseString
from datetime import datetime
import traceback
import pandas as pd
import numpy as np

#Variáveis globais
tabInfos = []

def obterdataHoraAtual():	
	return datetime.now()
	
def formataCampos(line):

	aux = []
	
	aux.append(str(line[0]).replace(' ', ''))
	aux.append("{:.9f}".format(line[1]).replace(' ', ''))
	aux.append("{:.9f}".format(line[2]).replace(' ', ''))
	aux.append("{:.9f}".format(line[3]).replace(' ', ''))
	aux.append(str(line[4]).replace(' ', ''))
	aux.append(str(line[5]).replace(' ', ''))
	
	tabInfos.append(aux)

def carregaDadosExcel():
	
	fileName = input('Por favor, digite o nome do arquivo excel: ')

	dataset = pd.read_excel(fileName, header=None, index=False).drop([0, 0])
	
	rows, cols = dataset.shape
	
	for row in range(1, rows+1):
		formataCampos(dataset.loc[row])

def estruturaNovoContrato(parseXML):

	elementRoot = parseXML.documentElement

	rootNodes = elementRoot.childNodes
	elementRoot.removeChild(rootNodes[0])
	elementRoot.removeChild(rootNodes[1])
	elementRoot.removeChild(rootNodes[2])

	lista = parseXML.getElementsByTagName('invocation')

	elementInvocation = lista[0].cloneNode(True)

	#Excluir primeiro Node invocation
	elementRoot.removeChild(rootNodes[1])


	last = elementInvocation.childNodes
	elementInvocation.removeChild(last[0])
	elementInvocation.removeChild(last[1])
	elementInvocation.removeChild(last[2])
	elementInvocation.removeChild(last[3])
	elementInvocation.removeChild(last[4])

	lista2 = last[0].childNodes
	last[0].removeChild(lista2[0])
	last[0].removeChild(lista2[1])
	last[0].removeChild(lista2[2])

	return elementInvocation

def insereNovosContratos(elementbase, elementRoot):

	invoiceDate = input('Por favor, digite a invoiceDate - aaaammdd: ')
	periodFrom = input('Por favor, digite o periodFrom - aaaammdd: ')
	
	dateTimeNow = obterdataHoraAtual()
	
	migrationDate = dateTimeNow.strftime('%Y-%m-%d') + 'T00:00:00'
	financeSystemTransferDate = dateTimeNow.strftime('%Y-%m-%d') + 'T12:00:00'
	dateTime = dateTimeNow.strftime('%Y-%m-%dT%H:%M:%S')

	countLinha = 0 
	for coluna in tabInfos:
		
		elementInvocation = elementbase.cloneNode(True)
        
		listElements = elementInvocation.childNodes

		listElements[0].setAttribute('migrationDate', migrationDate)
		listElements[0].setAttribute('invoiceDate', invoiceDate)
		listElements[0].setAttribute('calculationReferenceDate', invoiceDate)
		listElements[0].setAttribute('documentDate', invoiceDate)
		listElements[0].setAttribute('financeSystemTransferDate', financeSystemTransferDate)
		
		netAmount = coluna[1].replace(',','.')
		grossAmount = coluna[2].replace(',','.')
		vatAmount = coluna[3].replace(',','.')
		
		listElements[0].setAttribute('netAmount', netAmount)
		listElements[0].setAttribute('grossAmount', grossAmount)
		listElements[0].setAttribute('vatAmount', vatAmount)
		
		invoiceNumberFinance = '00000' + coluna[5]
		listElements[0].setAttribute('invoiceNumberFinance', invoiceNumberFinance)
		
		listContractNumber = coluna[0].split('/')
		
		if listContractNumber[0].isnumeric(): 
			contractNumber = str(int(listContractNumber[0]))
		else:
			contractNumber = str(listContractNumber[0])
			
		
		externalId = contractNumber + '-' + str(int(listContractNumber[1])) + '-' + invoiceNumberFinance
		
		listElements[0].setAttribute('externalId', externalId)
		listElements[0].setAttribute('sourceSystem', 'xmlFiles')
		listElements[0].setAttribute('customerInvoiceDefinition', 'MANU')
		
		
		listSubElements = listElements[0].childNodes
	
		listSubElements[0].setAttribute('number', coluna[4])
		listSubElements[1].setAttribute('creationDate', invoiceDate)
		listSubElements[1].setAttribute('periodFrom', periodFrom)
		listSubElements[1].setAttribute('invoicePositionNetAmount', netAmount)		
		listSubElements[1].setAttribute('invoicePositionVATAmount', vatAmount)

		#Removido conforme necessidade apontada pela Silvia
		#listElements[1].firstChild.data = contractNumber + '/' + str(int(listContractNumber[1]))
		listElements[1].firstChild.data = ''
			
		listElements[3].firstChild.data = coluna[0]
		
		elementRoot.appendChild(elementInvocation)
			

		countLinha += 1
	
def salvaXML(parseXML):

	elementExecutionSettings = parseXML.getElementsByTagName('executionSettings')
	
	dateTimeNow = obterdataHoraAtual()
	
	dateTime = dateTimeNow.strftime('%Y-%m-%dT%H:%M:%S')

	elementExecutionSettings[0].setAttribute('dateTime', dateTime)
	tenantId = elementExecutionSettings[0].getAttribute('tenantId')

	dataHoraAtual = dateTimeNow.strftime('%Y%m%dT%H%M%S')
	nameFile = '08-Revenue_' + dataHoraAtual + '_' + tenantId + '_ICON_tec_EDF01_Revenue_00001.xml'

	data = open(nameFile, 'w')
	parseXML.writexml(data, addindent="  ", newl="\n")

def mensagemErro():
	input("Ocorrem erros ao executar o programa, consulte o arquivo erros.txt para mais detalhes.")

	arq = open('erros.txt', 'w')

	var = traceback.format_exc()
	arq.write(var)
	arq.close()

def main():

	print("Bem vindo ao gerador                                  ")
	print("               __   ____  __ _                        ")
	print("               \ \ / /  \/  | |                       ")
	print("                \ V /| \  / | |                       ")
	print("                 > < | |\/| | |     	                 ")
	print("                / . \| |  | | |____                   ")
	print("               /_/ \_\_|  |_|______|     version 1.0.1")
	print("                                                      ")
	print("                       Developed by: Alexandre Santos ")
	print("                       E-mail: alexandre.s@daimler.com")
	print("                                                      ")
	
	carregaDadosExcel()
	
	datasource = open('modeloBase.xml')
	parseXML = parse(datasource)  # parse an open file

	elementRoot = parseXML.documentElement # get element root

	elementbase = estruturaNovoContrato(parseXML)

	insereNovosContratos(elementbase, elementRoot)

	salvaXML(parseXML)

	input("\nOperacão concluída com sucesso.\n\nPressione <enter> para sair")
	
if __name__ == "__main__": 
  
    # calling main function 
	try:
		main()
	except:
		mensagemErro()
	