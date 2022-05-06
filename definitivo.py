import xlsxwriter
import requests


dic = {'https://import-beltra.ozmap.com.br:9994/api/v2/users': 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJtb2R1bGUiOiJhcGkiLCJ1c2VyIjoiNWQ5ZjNmYjgyMDAxNDEwMDA2NDdmNzY4IiwiY3JlYXRpb25EYXRlIjoiMjAyMi0wMy0yNFQyMDo0NTowOC40MDZaIiwiaWF0IjoxNjQ4MTU0NzA4fQ.3Vg39IhsFa2fSywiqc3xGNrIu-ZTpmGSzxrQ00JJxsc',
    'https://import-beltra2.ozmap.com.br:9994/api/v2/users': 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJtb2R1bGUiOiJhcGkiLCJ1c2VyIjoiNWQ5ZjNmYjgyMDAxNDEwMDA2NDdmNzY4IiwiY3JlYXRpb25EYXRlIjoiMjAyMi0wMy0yNFQyMjoyNzowOC44NTdaIiwiaWF0IjoxNjQ4MTYwODI4fQ.Jce1yuBY-k6mK2ywpJlJb3VB5Tn_GCpVH1u7r_Yoxeg',
    'https://import-beltra3.ozmap.com.br:9994': 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJtb2R1bGUiOiJhcGkiLCJ1c2VyIjoiNjIyOGRmZmVkMmZlZTYwMDIwZjg0ZmViIiwiY3JlYXRpb25EYXRlIjoiMjAyMi0wMy0yNFQyMjoyNzozNS4zMjBaIiwiaWF0IjoxNjQ4MTYwODU1fQ.F6YuznPvc4sfmEFQCJMf2hMWa4lRxjmPo16UVDqaLC0',
    'https://teste.ozmap.com.br:9994/api/v2/users': 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJtb2R1bGUiOiJhcGkiLCJ1c2VyIjoiNWQ5ZjNmYjgyMDAxNDEwMDA2NDdmNzY4IiwiY3JlYXRpb25EYXRlIjoiMjAyMi0wNC0xM1QyMDoyNjowOS4wMzBaIiwiaWF0IjoxNjQ5ODgxNTY5fQ.8SCiUdS-vlpyz4MSeFjfd7ipbH97DklyMyTIKbuaMWo'}

class ApiConsummer:

    def __init__(self, response, token):
        headers = {'Authorization': 'Bearer ' + token}
        self.response = requests.get(response, headers=headers)

    def get_response(self):
        return self.response.json()

    def retorna_valores(self) -> dict:
        cont_OZmap, cont_OZmob, cont_Loki, cont_API = 0, 0, 0, 0

        numeros = {}
        for i in empresa['rows']:
            for j in i['resources']:
                if (j == 'OZmap'):
                    cont_OZmap += 1
                    numeros.update({j: cont_OZmap})
                if (j == 'OZmob'):
                    cont_OZmob += 1
                    numeros.update({j: cont_OZmob})
                if (j == 'Loki'):
                    cont_Loki += 1
                    numeros.update({j: cont_Loki})
                if (j == 'API'):
                    cont_API += 1
                    numeros.update({j: cont_API})

        return numeros


def CreateFile(lista):
        workbook = xlsxwriter.Workbook('usuarios_por_modulo.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.set_column('B:E', 18)
        format_1 = workbook.add_format({'bold': 1, 'italic': 1, 'align': 'center'})
        format_2 = workbook.add_format({'align': 'center'})

        empresas = ['EMPRESA A', 'EMPRESA B', 'EMPRESA C', 'EMPRESA D']
        modulos = ['OZmap', 'OzMob', 'Loki', 'API']

        count_empresas = 0
        for col in range(1,5):
            for row in range(1):
                worksheet.write(row, col, empresas[count_empresas], format_1)
                count_empresas += 1

        count_modulos = 0
        for row in range(1, 5):
            for col in range(1):
                worksheet.write(row, col, modulos[count_modulos], format_1)
                count_modulos += 1

        count = 0
        for col in range(1,5):
            for row in range(1, 5):
                worksheet.write(row, col,lista[count], format_2)
                count += 1

        workbook.close()


lista = []
for chave, valor in dic.items():

    try:
        response = ApiConsummer(chave, valor)
        empresa = response.get_response()
        result = response.retorna_valores()
        [lista.append(valor) for chave, valor in result.items()]

    except:
        print(f'Ocorreu um erro de requisição: getaddrinfo ENOTFOUND import-beltra3.ozmap.com.br')
        [lista.append('Erro de requisição!') for i in range (0,4)]


CreateFile(lista)