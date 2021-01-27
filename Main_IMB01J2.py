import win32com.client
from win32com.client import makepy
from pylab import *
from operator import itemgetter
import random
import os
import csv
import numpy
import statistics

class DSS(object):  # Classe DSS
    def __init__(self, dssFileName):

        # Create a new instance of the DSS
        sys.argv = ["makepy", "OpenDSSEngine.DSS"]
        makepy.main()
        self.dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

        # Start the DSS
        if self.dssObj.Start(0) == False:
            print("DSS Failed to Start")
        else:
            self.dssFileName = dssFileName
            # Assign a variable to each of the interfaces for easier access
            self.dssText = self.dssObj.Text
            self.dssCircuit = self.dssObj.ActiveCircuit
            self.dssSolution = self.dssCircuit.Solution
            self.dssCktElement = self.dssCircuit.ActiveCktElement
            self.dssBus = self.dssCircuit.ActiveBus
            self.dssTransformer = self.dssCircuit.Transformers

    def compile_DSS(self):
        # Always a good idea to clear the DSS when loading a new circuit
        # self.dssObj.ClearAll()

        # Load the given circuit master file into OpenDSS
        self.dssText.Command = "compile " + self.dssFileName

        # OpenDSS folder
        self.OpenDSS_folder_path = os.path.dirname(self.dssFileName)

    def solve(self, solucao):
        # self.compile_DSS()
        self.results_path = self.OpenDSS_folder_path + "/results_Main"
        self.dssText.Command = "set DataPath=" + self.results_path

        # Monitores
        listaTrafosDist = self.listaTrafos()
        for i in listaTrafosDist:
            self.dssText.Command = "New Monitor." + str(listaTrafosDist.index(i)) + " Element=" + i + " mode=32 terminal=2 "

        kWRatedList = list(range(100, 3100, 100))
        # PmppList = list(range(100, 4700, 200))
        LoadshapePointsList = [round(ctd, 2) for ctd in list(numpy.arange(-1.0, 1.05, 0.05))]
        Loadshape = [LoadshapePointsList[ctd] for ctd in solucao[:]]
        Loadshape = self.LoadshapeToMediaMovel(Loadshape)
        # print(Loadshape)

        self.dssText.Command = "Loadshape.Loadshape1.mult=" + str(Loadshape)
        self.dssText.Command = "Storage.storage.Bus1=" + '107139M3009 '
        self.dssText.Command = "PVSystem.PV.Bus1=" + '107139M3009 '
        # self.dssText.Command = "Storage.storage.kWrated=" + str(kWRatedList[solucao[0]])
        # self.dssText.Command = "Storage.storage.kva=" + str(kWRatedList[solucao[0]])
        # self.dssText.Command = "Storage.storage.kw=" + str(kWRatedList[solucao[0]])
        self.dssText.Command = "Storage.storage.kWrated=1000"
        self.dssText.Command = "Storage.storage.kva=1000"
        self.dssText.Command = "Storage.storage.kw=1000"
        self.dssText.Command = "PVSystem.PV.KVA=" + '2500'
        self.dssText.Command = "PVSystem.PV.Pmpp=" + '2500'
        self.dssText.Command = "Storage.storage.enabled=yes"

        self.dssSolution.Solve()

        self.dssText.Command = "export meters"
        self.dssText.Command = "export monitor Potencia_Feeder"

        for i in listaTrafosDist:
            self.dssText.Command = "export monitor " + str(listaTrafosDist.index(i))

    def funcaoCusto(self, solucao):
        # d = DSS(r"D:\UFBA\IC-storage\Algoritmo_Genetico\Main_ModoFollow_Trafo.F21898.dss")
        self.compile_DSS()
        self.solve(solucao)

        # Inclinaçoes
        Inclinacao = 0
        ListaInclinacoes = self.InclinacoesLoadshape(solucao)

        for i in ListaInclinacoes:
            if numpy.abs(i) > 40:
                Inclinacao += numpy.abs(i)

        # Punição Niveis de Tensão
        if self.BarrasTensaoVioladas() > self.BarrasTensaoVioladasOriginal:
            PunicaoTensao = 9999999999
        else:
            PunicaoTensao = 0

        # PESOS
        a = 0.5  # Perdas
        b = 0.5  # DP do Carregamento do trafo

        # PERDAS
        ### Acessando arquivo CSV Potência
        dataEnergymeterCSV = {}
        self.dataperda = {}

        fname = "D:\\UFBA/IC-storage\\AG_IMB01J2\\IMB01J2\\results_Main\\CABIMB_EXP_METERS.csv"

        with open(str(fname), 'r', newline='') as file:
            csv_reader_object = csv.reader(file)
            name_col = next(csv_reader_object)

            for row in name_col:
                dataEnergymeterCSV[row] = []

            for row in csv_reader_object:  ##Varendo todas as linhas
                for ndata in range(0, len(name_col)):  ## Varendo todas as colunas
                    rowdata = row[ndata].replace(" ", "").replace('"',"")
                    if rowdata == "FEEDER" or rowdata == "":
                        dataEnergymeterCSV[name_col[ndata]].append(rowdata)
                    else:
                        dataEnergymeterCSV[name_col[ndata]].append(float(rowdata))

        self.dataperda['Perdas %'] = (dataEnergymeterCSV[' "Zone Losses kWh"'][0]/dataEnergymeterCSV[' "Zone kWh"'][0])*100
        os.remove(fname)

        # DESVIO PADRÃO DO CARREGAMENTO DO TRAFO
        ### Acessando arquivo CSV Potência
        dataFeederMmonitorCSV = {}

        fname = "D:\\UFBA/IC-storage\\AG_IMB01J2\\IMB01J2\\results_Main\\CABIMB_Mon_potencia_feeder_1.csv"

        with open(str(fname), 'r', newline='') as file:
            csv_reader_object = csv.reader(file)
            name_col = next(csv_reader_object)

            for row in name_col:
                dataFeederMmonitorCSV[row] = []

            dataFeederMmonitorCSV['PTotal'] = []

            for row in csv_reader_object:  ##Varendo todas as linhas
                Pt = 0
                for ndata in range(0, len(name_col)):  ## Varendo todas as colunas
                    rowdata = row[ndata].replace(" ", "").replace('"',"")
                    dataFeederMmonitorCSV[name_col[ndata]].append(float(rowdata))
                    if name_col[ndata] == ' P1 (kW)' or name_col[ndata] == ' P2 (kW)' or name_col[ndata] == ' P3 (kW)':
                        Pt += float(rowdata)

                dataFeederMmonitorCSV['PTotal'].append(Pt)
        Desvio = statistics.pstdev(dataFeederMmonitorCSV['PTotal'])
        Perdas_sem_Pv_Stor = 2.316
        Custo = a/(Perdas_sem_Pv_Stor/100-self.dataperda['Perdas %']/100) + b*Desvio + Inclinacao + PunicaoTensao
        # Custo = 1/self.dataperda['Perdas %']/100 + b*Desvio + PunicaoTensao
        return Custo

    def mutacao(self, dominio, passo, solucao):
        i = random.randint(0, len(dominio) - 1)
        mutante = solucao
        # print("mutacao", mutante, i)
        if random.random() < 0.5:
            if solucao[i] != dominio[i][0]:
                if solucao[i] >= (dominio[i][0] + passo):
                    mutante = solucao[0:i] + [solucao[i] - passo] + solucao[i + 1:]
        else:
            if solucao[i] != dominio[i][1]:
                if solucao[i] <= (dominio[i][1] - passo):
                    mutante = solucao[0:i] + [solucao[i] + passo] + solucao[i + 1:]

        return mutante

    def cruzamento(self, dominio, individuo1, individuo2):
        i = random.randint(1, len(dominio) - 2)
        return individuo1[0:i] + individuo2[i:]

    def genetico(self, dominio, tamanho_populacao=80,  passo=1,
                 probabilidade_mutacao=0.2, elitismo=0.2, numero_geracoes=300):

        self.compile_DSS()
        self.results_path = self.OpenDSS_folder_path + "/results_Main"
        self.dssText.Command = "set DataPath=" + self.results_path

        # Monitores
        listaTrafosDist = self.listaTrafos()
        for i in listaTrafosDist:
            self.dssText.Command = "New Monitor." + str(listaTrafosDist.index(i)) + " Element=" + i + " mode=32 terminal=2 "

        self.dssText.Command = "Storage.storage.enabled=no"
        self.dssText.Command = "PVSystem.PV.enabled=no"
        self.dssSolution.Solve()

        for i in listaTrafosDist:
            self.dssText.Command = "export monitor " + str(listaTrafosDist.index(i))

        self.BarrasTensaoVioladasOriginal = self.BarrasTensaoVioladas()

        print(self.BarrasTensaoVioladasOriginal)

        populacao = []
        kWRatedList = list(range(100, 3100, 100))
        LoadshapePointsList = [round(ctd, 2) for ctd in list(numpy.arange(-1.0, 1.05, 0.05))]
        listadeLoadShapes1 = [
            [0, 0, -0.3, -0.45, -0.5, -0.45, -0.3, 0, 0, 0, 0, 0, 0, 0, 0, 0.3, 0.5, 0.8, 0.9, 0.8, 0.5, 0.3, 0, 0],
            [0, 0, -0.3, -0.45, -0.5, -0.45, -0.3, 0, 0, 0, 0, 0, 0, 0, 0, 0.3, 0.4, 0.6, 0.8, 0.9, 0.8, 0.5, 0.3, 0],
            [0, 0, 0, -0.3, -0.45, -0.5, -0.45, -0.3, 0, 0, 0, 0, 0, 0, 0, 0, 0.3, 0.6, 0.75, 0.95, 0.9, 0.8, 0.3, 0],
            [0, -0.1, -0.2, -0.2, -0.2, -0.2, -0.2, -0.2, -0.1, 0, 0, 0, 0, 0, 0.3, 0.6, 0.8, 0.8, 0.8, 0.6, 0.4, 0, 0,
             0],
            [0.3, 0.3, 0.3, 0.3, 0.3, 0.2, 0.1, 0, 0, -0.5, -0.6, -0.7, -0.8, -0.9, -0.9, -0.8, -0.4, 0.3, 0.5, 0.8,
             0.9, 0.7, 0.3, 0.3],
            [0, -0.1, -0.2, -0.2, 0, 0, 0, 0, -0.1, -0.3, -0.65, -0.7, -0.8, -0.9, -0.85, -0.75, -0.45, 0.5, 0.9, 0.9,
             0.95, 0.8, 0.8, 0.7],
            [0.3, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0, -0.1, -0.3, -0.6, -0.75, -0.75, -0.8, -0.9, -0.85, -0.4, 0.5, 0.9,
             0.9, 0.9, 0.8, 0.8, 0.7],
            [0.3, 0.3, 0.5, 0.5, 0.5, 0.3, 0.3, 0.3, 0, 0, -0.8, -0.8, -0.9, -0.9, -0.8, -0.8, -0.8, 0, 0.75, 0.8, 0.9,
             0.8, 0.8, 0.4],
            [0.2, 0.25, 0.15, 0.2, 0.2, 0.2, 0.65, 0.7, 0.7, -0.3, -0.65, -0.65, -0.75, -0.85, -0.95, -0.95, -0.45,
             0.45, 0.85, 0.85, 0.85, 0.85, 0.7, 0.35],
            [0.15, 0.15, 0.15, 0, 0, 0, 0, 0, -0.05, -0.05, -0.45, -0.45, -0.75, -0.75, -0.75, -0.75, -0.75, -0.75, 0.3,
             0.45, 0.55, 0.5, 0.5, 0.05],
            [0.05, 0, 0, 0, 0, 0, 0, 0, -0.1, -0.25, -0.35, -0.55, -0.6, -0.9, -0.9, -0.95, -0.5, 0.45, 0.75, 0.85, 0.9,
             0.9, 0.75, 0.15],
            [0.05, 0.45, 0.75, 0.75, 0.75, 0.75, 0.75, 0.15, 0, 0, -0.35, -0.35, -0.95, -0.95, -0.95, -0.95, -0.4, 0.45,
             0.55, 0.6, 0.6, 0.6, 0.6, 0.15],
            [0.1, -0.2, -0.2, -0.2, -0.2, 0, 0, 0, -0.1, -0.25, -0.4, -0.6, -0.6, -0.6, -0.6, -0.5, -0.4, 0.6, 0.8, 0.9,
             0.8, 0.5, 0.3, 0],
            [-0.15, -0.15, -0.15, -0.15, 0, 0, 0, 0, 0, -0.2, -0.4, -0.5, -0.7, -0.7, -0.5, -0.4, -0.2, 0.45, 0.75, 0.8,
             0.75, 0.75, 0.5, -0.25],
            [0.2, 0.25, 0.2, 0.25, 0.2, 0.25, 0.2, 0.25, -0.2, -0.25, -0.6, -0.6, -0.7, -0.8, -0.8, -0.8, -0.6, 0.45,
             0.9, 0.95, 0.95, 0.9, 0.75, 0.1],
            [0, 0.05, 0.1, 0.15, 0.15, 0, 0, 0, -0.1, -0.3, -0.4, -0.6, -0.7, -0.75, -0.85, -0.8, -0.8, 0.25, 0.5, 0.7,
             0.7, 0.9, 0.3, 0.1],
            [-0.3, -0.3, -0.3, -0.3, -0.3, -0.05, -0.05, -0.05, -0.1, -0.15, -0.3, -0.3, -0.5, -0.55, -0.55, -0.65,
             -0.5, 0.25, 0.5, 0.9, 0.9, 0.9, 0.5, 0.4],
            [-0.25, -0.3, -0.25, -0.3, -0.25, -0.3, 0, 0, 0, -0.35, -0.55, -0.65, -0.7, -0.75, -0.95, -0.4, -0.3, 0.2,
             0.5, 0.85, 0.8, 0.85, 0.5, 0.4],
            [0.1, 0.1, 0.1, 0.1, 0.2, 0.35, 0.35, 0.45, 0.1, -0.3, -0.5, -0.7, -0.75, -0.85, -0.85, -0.85, -0.75, -0.15,
             0.7, 0.75, 0.75, 0.75, 0.3, 0.25],
            [0.3, 0.3, 0.3, 0.3, 0.3, 0.45, 0.45, 0.45, 0.15, -0.2, -0.5, -0.65, -0.75, -0.8, -0.85, -0.85, -0.8, -0.25,
             0.5, 0.8, 0.85, 0.85, 0.5, 0.2],
            [0.3, 0.3, 0.3, 0.3, 0.3, 0.45, 0.45, 0.45, 0.15, -0.3, -0.4, -0.6, -0.7, -0.75, -0.85, -0.8, -0.8, 0.25,
             0.5, 0.9, 0.9, 0.9, 0.6, 0.4],
            [0.35, 0.35, 0, 0, 0, 0.4, 0.45, 0.45, 0.15, -0.3, -0.55, -0.65, -0.7, -0.75, -0.85, -0.8, -0.8, 0.25, 0.5,
             0.9, 0.9, 0.9, 0.6, 0.4],
            [0.3, 0.15, 0.15, 0.15, 0.3, 0.4, 0.45, 0.45, 0.15, -0.3, -0.4, -0.6, -0.75, -0.75, -0.85, -0.8, -0.8, 0.25,
             0.45, 0.8, 0.9, 0.95, 0.6, 0.4],
            [0.3, 0.2, 0, 0, 0.1, 0.15, 0.15, 0.15, 0.15, -0.3, -0.4, -0.6, -0.7, -0.95, -0.95, -0.8, -0.8, 0.25, 0.45,
             0.9, 0.9, 0.9, 0.6, 0.4],
            [0.1, 0.15, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, -0.7, -0.7, -0.85, -0.85, -0.85, -0.3, 0.25, 0.5,
             0.6, 0.8, 0.85, 0.7, 0.4],
            [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.5, -0.5, -0.5, -0.5, -0.5, -0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0]
        ]
        listadeLoadShapes2 = [
            [-0.3,-0.3,-0.3,-0.3,-0.3,-0.3,-0.3,0,0.3,0.45,0.5,0.5,0.45,0.3,0.1,-0.1,-0.1,-0.1,-0.1,-0.1,-0.1,-0.1,-0.1,-0.1],
            [-0.3,-0.45,-0.5,-0.45,-0.3,0,0,0,0,0,0,0,0,0,0,0.3,0.4,0.6,0.8,0.9,0.8,0.5,0.3,0],
            [0,0,0,-0.3,-0.45,-0.5,-0.45,-0.3,0,0,0,0,0,0,0,0,0.3,0.6,0.75,0.95,0.9,0.8,0.3,0],
            [0,0,0,0,0,0,0,0,0,0,0,-0.5,-0.5,-0.5,-0.5,-0.5,-0.5,0.5,0.5,0.5,0.5,0.5,0.5,0],
            [0.1,0.1,0.1,0.1,0.1,0.1,0.1,0,-0.1,-0.3,-0.6,-0.75,-0.75,-0.8,-0.9,-0.85,-0.4,0.5,0.9,0.9,0.9,0.8,0.8,0.7],
            [0,0.3,0.4,0.6,0.7,0.6,0.4,0.3,0,-0.3,-0.45,-0.5,-0.45,-0.4,-0.3,-0.05,0.25,0.4,0.55,0.6,0.6,0.45,0.25,0.15],
            [0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0.4,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0],
            [0.05,0.05,0.05,0.05,0.05,0.05,0.05,0.05,0.05,-0.3,-0.3,-0.3,-0.3,-0.3,-0.3,-0.3,-0.3,0.05,0.05,0.05,0.05,0.05,0.05,0.05],
            [0.15,0.15,0.15,0.15,0.15,0.15,0.15,0.15,0.15,0.15,0.15,0,0,0,0,0,0,-0.35,-0.35,-0.35,-0.35,-0.35,-0.35,-0.35],
            [-0.2,-0.2,-0.2,-0.2,-0.2,-0.2,-0.2,-0.2,-0.2,-0.2,0.2,0.2,0.2,0.2,0.2,0.2,0.2,0.2,0.4,0.4,0.4,0.4,0.4,0.4]
        ]

        for i in range(tamanho_populacao):
            # Solucao para todos os valores Random
            solucao = []
            for ctd in range(len(dominio)):
                if ctd == 0:
                    solucao.append(random.randint(dominio[ctd][0], dominio[ctd][1]))
                else:
                    a = [dominio[ctd][0], solucao[-1] - 14]
                    a = max(a)
                    b = [dominio[ctd][1], solucao[-1] + 14]
                    b = min(b)
                    solucao.append(random.randint(a, b))
            # solucao = [random.randint(dominio[i][0], dominio[i][1]) for i in range(len(dominio))]
            # print(solucao)
            populacao.append(solucao)

        numero_elitismo = int(elitismo * tamanho_populacao)
        geracao = 1

        for i in range(numero_geracoes):
            start = time.time()
            custos = [(self.funcaoCusto(individuo), individuo) for individuo in populacao]
            custos.sort()
            # custos_traduzidos = [(ctd[0], kWRatedList[ctd[1][0]], [LoadshapePointsList[i] for i in ctd[1][1:]]) for ctd in custos]
            custos_traduzidos = [(ctd[0], [LoadshapePointsList[i] for i in ctd[1][:]]) for ctd in custos]
            print("Geração", geracao,  custos_traduzidos)
            # print('custos', custos)
            self.CalculaCustos(custos[0][1])
            geracao += 1
            individuos_ordenados = [individuo for (custo, individuo) in custos]
            populacao = individuos_ordenados[0:numero_elitismo]
            lista_rank = [(individuo, (tamanho_populacao - individuos_ordenados.index(individuo))/(tamanho_populacao*(tamanho_populacao-1))) for individuo in individuos_ordenados]
            lista_rank.reverse()
            # print("lista_rank", lista_rank)
            soma=0
            for ctd in lista_rank:
                soma += ctd[1]

            while len(populacao) < tamanho_populacao:
                if random.random() < probabilidade_mutacao:
                    m = random.randint(0, numero_elitismo)
                    populacao.append(self.mutacao(dominio, passo, individuos_ordenados[m]))
                else:
                    aleatorio = random.uniform(0, soma)
                    # print('aleatorio', aleatorio)
                    s = 0
                    for j in lista_rank:
                        s += j[1]
                        if aleatorio < s:
                            c1 = j[0]
                            # print('c1', c1)
                            break
                    aleatorio = random.uniform(0, soma)
                    s = 0
                    for j in lista_rank:
                        s += j[1]
                        if aleatorio < s:
                            c2 = j[0]
                            # print('c2', c2)
                            break
                    populacao.append(self.cruzamento(dominio, c1, c2))
                    # c1 = random.randint(0, numero_elitismo)
                    # c2 = random.randint(0, numero_elitismo)
                    # populacao.append(self.cruzamento(dominio, individuos_ordenados[c1], individuos_ordenados[c2]))

            end = time.time()
            print("Tempo da geração:", end - start)
        return custos[0][1]

    def BarrasTensaoVioladas(self):
        BarrasVioladas = 0
        listaTrafosDist = self.listaTrafos()

        for i in listaTrafosDist:
            dataMonitorTrafoDist = {}
            fname = "D:\\UFBA/IC-storage\\AG_IMB01J2\\IMB01J2\\results_Main\\CABIMB_Mon_" + str(listaTrafosDist.index(i)) + "_1.csv"

            with open(str(fname), 'r', newline='') as file:
                csv_reader_object = csv.reader(file)
                name_col = next(csv_reader_object)

                for row in name_col:
                    dataMonitorTrafoDist[row] = []

                for row in csv_reader_object:  ##Varendo todas as linhas
                    for ndata in range(0, len(name_col)):  ## Varendo todas as colunas
                        rowdata = row[ndata].replace(" ", "").replace('"',"")
                        if name_col[ndata] == ' |V|1 (volts)' or name_col[ndata] == ' |V|2 (volts)' or name_col[ndata] == ' |V|3 (volts)':
                            dataMonitorTrafoDist[name_col[ndata]].append(float(rowdata)/127)

            TensaoPUFasesBarras = dataMonitorTrafoDist[' |V|1 (volts)'] + dataMonitorTrafoDist[' |V|2 (volts)'] + dataMonitorTrafoDist[' |V|3 (volts)']
            # print(TensaoPUFasesBarras)
            for ctd in TensaoPUFasesBarras:
                if ctd > 1.03 or ctd < 0.97:
                    BarrasVioladas += 1

        # TensaoPUFasesBarras = d.dssCircuit.AllNodeVmagPUByPhase(1) + d.dssCircuit.AllNodeVmagPUByPhase(2) + d.dssCircuit.AllNodeVmagPUByPhase(3)
        # for i in TensaoPUFasesBarras:
        #     if i > 1.03 or i < 0.97:
        #         BarrasVioladas += 1
        return BarrasVioladas

    def InclinacoesLoadshape(self, solucao):
        LoadshapePointsList = [round(ctd, 2) for ctd in list(numpy.arange(-1.0, 1.05, 0.05))]
        Loadshape = [LoadshapePointsList[i] for i in solucao]
        Inclinacoes = []

        for i in range((len(Loadshape)-1)):
            x = Loadshape[i+1] - Loadshape[i]
            Inclinacoes.append(numpy.arctan(x)*180/pi)

        return Inclinacoes

    def CalculaCustos(self, solucao):
        # d = DSS(r"D:\UFBA\IC-storage\Algoritmo_Genetico\Main_ModoFollow_Trafo.F21898.dss")
        self.compile_DSS()
        self.solve(solucao)

        LoadshapePointsList = [round(ctd, 2) for ctd in list(numpy.arange(-1.0, 1.05, 0.05))]
        Loadshape = [LoadshapePointsList[ctd] for ctd in solucao[:]]

        # Inclinaçoes
        Inclinacao = 0
        ListaInclinacoes = self.InclinacoesLoadshape(solucao)

        for i in ListaInclinacoes:
            if numpy.abs(i) > 40:
                Inclinacao += numpy.abs(i)

        ### Acessando arquivo CSV Potência
        dataEnergymeterCSV = {}
        self.dataperda = {}

        fname = "D:\\UFBA/IC-storage\\AG_IMB01J2\\IMB01J2\\results_Main\\CABIMB_EXP_METERS.csv"


        with open(str(fname), 'r', newline='') as file:
            csv_reader_object = csv.reader(file)
            name_col = next(csv_reader_object)

            for row in name_col:
                dataEnergymeterCSV[row] = []

            for row in csv_reader_object:  ##Varendo todas as linhas
                for ndata in range(0, len(name_col)):  ## Varendo todas as colunas
                    rowdata = row[ndata].replace(" ", "").replace('"',"")
                    if rowdata == "FEEDER" or rowdata == "":
                        dataEnergymeterCSV[name_col[ndata]].append(rowdata)
                    else:
                        dataEnergymeterCSV[name_col[ndata]].append(float(rowdata))

        self.dataperda['Perdas %'] = (dataEnergymeterCSV[' "Zone Losses kWh"'][0]/dataEnergymeterCSV[' "Zone kWh"'][0])*100
        os.remove(fname)

        ### Acessando arquivo CSV Potência
        dataFeederMmonitorCSV = {}

        fname = "D:\\UFBA/IC-storage\\AG_IMB01J2\\IMB01J2\\results_Main\\CABIMB_Mon_potencia_feeder_1.csv"

        with open(str(fname), 'r', newline='') as file:
            csv_reader_object = csv.reader(file)
            name_col = next(csv_reader_object)

            for row in name_col:
                dataFeederMmonitorCSV[row] = []

            dataFeederMmonitorCSV['PTotal'] = []

            for row in csv_reader_object:  ##Varendo todas as linhas
                Pt = 0
                for ndata in range(0, len(name_col)):  ## Varendo todas as colunas
                    rowdata = row[ndata].replace(" ", "").replace('"', "")
                    dataFeederMmonitorCSV[name_col[ndata]].append(float(rowdata))
                    if name_col[ndata] == ' P1 (kW)' or name_col[ndata] == ' P2 (kW)' or name_col[ndata] == ' P3 (kW)':
                        Pt += float(rowdata)

                dataFeederMmonitorCSV['PTotal'].append(Pt)

        print('Perdas:', self.dataperda['Perdas %'], 'Inclinação:', Inclinacao, 'Barras_Violada:', self.BarrasTensaoVioladas(), 'PTotal:', dataFeederMmonitorCSV['PTotal'])
        print('MM:', self.LoadshapeToMediaMovel(Loadshape))
        return self.dataperda['Perdas %']

    def listaTrafos(self):
        dataTrafoDistDSS = []
        for linha in open('D:\\UFBA/IC-storage\\AG_IMB01J2\\IMB01J2\\TrafoDist.dss'):
            if linha.split(" ")[0] != "!":
                dataTrafoDistDSS.append(linha.split(" ")[1])
        return dataTrafoDistDSS

    def LoadshapeToMediaMovel(self, solucao):
        medias_moveis = []
        num_media = 2
        i = 0
        while i < (len(solucao) - num_media + 1):
            grupo = solucao[i: i + num_media]
            media_grupo = sum(grupo) / num_media
            medias_moveis.append(media_grupo)
            i += 1
        medias_moveis.insert(0, medias_moveis[0])
        return medias_moveis

if __name__ == '__main__':
    d = DSS(r"D:\\UFBA/IC-storage\\AG_IMB01J2\\IMB01J2\\MAIN_IMB01J2.dss")
    kWRatedList = list(range(100, 3100, 100))
    # dominio = [(0, len(kWRatedList) - 1), (0, 40) , (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40)]
    # dominio para valores totalmente Random
    dominio = [(0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40)]
    # dominio para 0:8 <= 0; 9:16 >= 0; 17:23 <=0   - + -
    # dominio = [(0, 20) , (0, 20),  (0, 20),  (0, 20),  (0, 20),  (0, 20),  (0, 20),  (0, 20),  (0, 20),  (20, 40),  (20, 40),  (20, 40),  (20, 40),  (20, 40),  (20, 40),  (20, 40),  (20, 40),  (0, 20),  (0, 20),  (0, 20),  (0, 20),  (0, 20),  (0, 20),  (0, 20)]
    # dominio para 0:8 >= 0; 9:16 <= 0; 17:23 >=0   + - +
    #dominio = [(20, 40) , (20, 40),  (20, 40),  (20, 40),  (20, 40),  (20, 40),  (20, 40),  (20, 40),  (20, 40),  (0, 20),  (0, 20),  (0, 20),  (0, 20),  (0, 20),  (0, 20),  (0, 20),  (0, 20),  (20, 40),  (20, 40),  (20, 40),  (20, 40),  (20, 40),  (20, 40),  (20, 40)]

    # d.listaTrafos()
    solucao_genetico = d.genetico(dominio)
    custo_genetico = d.funcaoCusto(solucao_genetico)
    print(custo_genetico)
    print(solucao_genetico)
