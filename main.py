import os
import time
import asyncio
import aiohttp
from openpyxl import load_workbook

PASTA_DOWNLOADS = 'downloads'


def gerar_url(mlb):
    return f'https://api.mercadolibre.com/items/{mlb}'


async def pegar_titulo(mlb):
    url = gerar_url(mlb)

    async with aiohttp.ClientSession() as session:
        res = await session.get(url)
        if res.status == 200:
            print(f'{mlb} - OK')
            json = await res.json()
            await salvar_titulo(mlb, json['title'])


async def salvar_titulo(mlb, descricao):
    with open(mlb + '.txt', 'w') as arquivo:
        arquivo.write(str(descricao))


def gerar_lista_mlbs(nome_planilha):
    arquivo = load_workbook(nome_planilha)
    planilha = arquivo.active

    lista_mlbs = []

    for row in planilha.rows:
        for cell in row:
            lista_mlbs.append(cell.value)

    return lista_mlbs


def pegar_e_salvar_titulos():
    lista_mlbs = gerar_lista_mlbs('mlbs.xlsx')
    os.chdir(PASTA_DOWNLOADS)

    loop = asyncio.get_event_loop()
    tarefas = [pegar_titulo(mlb) for mlb in lista_mlbs]
    wait_coro = asyncio.wait(tarefas)
    res, _ = loop.run_until_complete(wait_coro)
    loop.close()

    return len(res)


if __name__ == '__main__':
    tempo_inicio = time.time()

    count = pegar_e_salvar_titulos()

    tempo_final = time.time() - tempo_inicio
    print(f'{count} titulos salvos em {round(tempo_final, 2)} segundos.')
