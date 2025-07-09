import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


def converter_resultados():
    """
    Lê uma planilha de resultados da loteria,
    formata os dados e salva em um arquivo de texto.
    """
    try:
        # cria uma janela raiz escondida, para poder mostrar as caixas de diálogo
        root = tk.Tk()
        root.withdraw()

        # SELECIONAR ARQUIVO DE ENTRADA
        messagebox.showinfo("Conversor de resultados de loteria",
                            "Selecione sua planilha Excel (.xlsx) com os resultados.")

        input_path = filedialog.askopenfilename(
            title="Selecione a planilha de entrada",
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )

        # se o usuário cancelar a seleção
        if not input_path:
            print("Nenhuma planilha selecionada. O programa será encerrado.")
            return

        print(f"Lendo o arquivo: {input_path}")

        # LER E PROCESSAR A PLANILHA

        # colunas de interesse (A: concurso, C - Q: bolas lotomania).
        # Para adaptar a outros concursos, só editar as colunas, com o número de bolas
        # começando na letra C. No caso, lotofácil são 15 letras, começando em C.
        # Coluna A = 0. Coluna C = 2.
        # LOTOFACIL = 15 bolas
        # LOTOMANIA = 20 bolas

        tipo = int(input('Digite o número de bolas'))
        colunasLer = [0] + list(range(2,2+tipo))
        # 2+tipo: começando pela coluna 2 (C), lê a quantidade de colunas de acordo
        # com a quantidade de bolas

        # le a planilha, sendo a primeira linha cabeçalho, lendo só as colunas necessarias
        df = pd.read_excel(input_path, header=None, usecols=colunasLer, dtype=object)

        linhasFormatadas = []

        # nomes reais das colunas lidas do arquivo
        colunaConcurso = df.columns[0]
        colunasBolas = df.columns[1:]

        print("Processando os concursos...")

        # loop pra cada linha da planilha
        for indice, linha in df.iterrows():
            try:
                # Ignora a primeira linha se ela for um cabeçalho
                if indice == 0 and not str(linha[colunaConcurso]).isdigit():
                    print("Cabeçalho detectado e ignorado.")
                    continue
                # formata em 4 digitos o numConcurso
                numConcurso = int(linha[colunaConcurso])
                concursoFormatado = f"{numConcurso:04d}"

                # formata em 2 digitos os numsBolas e coloca um espaço
                numsBolas = [int(linha[col]) for col in colunasBolas]
                bolasFormatadas = [f"{bola:02d}" for bola in numsBolas]
                bolasString = " ".join(bolasFormatadas)

                linhaFinal = f"{concursoFormatado}={bolasString}"
                linhasFormatadas.append(linhaFinal)

            except (ValueError, TypeError):
                # se encontrar um valor que não é um número
                print(f"Aviso: A linha {indice + 2} da planilha contém dados inválidos e foi ignorada.")
                continue

        if not linhasFormatadas:
            messagebox.showerror("Erro",
                                 "Nenhum dado válido foi encontrado na planilha. Verifique o formato do arquivo.")
            return

        print("Processamento concluído com sucesso!")

        # SALVAR O ARQUIVO DE SAÍDA
        messagebox.showinfo("Processamento Concluído", "Escolha onde salvar o arquivo final. Faça uma cópia do arquivo original ResultadosLF.con.")

        output_path = filedialog.asksaveasfilename(
            title="Salvar arquivo de resultado",
            defaultextension=".con",
            filetypes=[("Arquivos de Concursos", "*.con"), ("Todos os arquivos", "*.*")]
        )

        # se o usuário cancelar a tela de salvar
        if not output_path:
            print("Nenhum local de salvamento escolhido. O programa será encerrado.")
            return

        # Escreve todas as linhas formatadas no arquivo de texto
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(linhasFormatadas))

        messagebox.showinfo("Sucesso!", f"O arquivo foi salvo com sucesso em:\n{output_path}")
        print(f"Resultado salvo em: {output_path}")

    except Exception as e:
        # qualquer outro erro que possa acontecer
        messagebox.showerror("Erro Inesperado", f"Ocorreu um erro durante a execução:\n\n{e}")
        print(f"ERRO: {e}")


if __name__ == "__main__":
    converter_resultados()