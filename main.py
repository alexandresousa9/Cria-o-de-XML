#!/usr/bin/python
# -*- coding: UTF-8 -*-

from xml.dom.minidom import parse, parseString
import csv
import os

tab_infos = []

def carregaDadosCSV():
	path = os.path.dirname(os.path.abspath(__file__))
	arquivo = open(path + '/createXML.csv')
	linhas = csv.reader(arquivo, delimiter=';')
	
	for linha in linhas:
		tab_infos.append(linha)

def preparaXML(parseXML):

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

def incluiElementsXML(elementbase, elementRoot):

	countLinha = 0 
	for linha in tab_infos:
		
		elementInvocation = elementbase.cloneNode(True)
        
		listElements = elementInvocation.childNodes

		listElements[0].setAttribute('migrationDate', linha[0])
		listElements[0].setAttribute('InvoiceData', linha[1])
		listElements[0].setAttribute('calculationReferenceDate', linha[2])
		listElements[0].setAttribute('documentDate', linha[3])
		listElements[0].setAttribute('financeSystemTransferDate', linha[4])
		listElements[0].setAttribute('netAmout', linha[8])
		listElements[0].setAttribute('grossAmout', linha[9])
		listElements[0].setAttribute('invoiceNumberFinance', linha[10])
		listElements[0].setAttribute('externalId', linha[11])
		listElements[0].setAttribute('customerInvoiceDefinition', linha[15])
		
		listSubElements = listElements[0].childNodes
	
		listSubElements[0].setAttribute('externalId', linha[7])
		listSubElements[1].setAttribute('creationDate', linha[5])
		listSubElements[1].setAttribute('periodFrom', linha[6])
		listSubElements[1].setAttribute('invoicePositionNetAmount', linha[12])


		listElements[1].firstChild.data = linha[13]
		listElements[3].firstChild.data = linha[14]
		
		elementRoot.appendChild(elementInvocation)
			

		countLinha += 1
	


carregaDadosCSV()

datasource = open('teste.xml')
parseXML = parse(datasource)  # parse an open file
print(parseXML)

elementRoot = parseXML.documentElement # get element root

elementbase = preparaXML(parseXML)

incluiElementsXML(elementbase, elementRoot)

data = open('arquivoFinal.xml', 'w')
parseXML.writexml(data, addindent="  ", newl="\n")
