import pandas as pd
import random
import string
import tkinter as tk
from tkinter import filedialog, messagebox

def gerar_id_projeto(codigo_projeto):
    """Gera um ID com base no código do projeto fornecido."""
    return codigo_projeto.replace('-', '') + 'v1'

def excel_para_js(caminho_excel, caminho_js, id_projeto):
    # Tenta ler a planilha Excel
    try:
        df = pd.read_excel(caminho_excel)
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível ler o arquivo Excel: {e}")
        return

    # Gera o arquivo .js
    try:
        with open(caminho_js, 'w', encoding='utf-8') as arquivo_js:
            # Cabeçalho fixo do JS
            arquivo_js.write('var conf_mp = {\n')
            arquivo_js.write(f'    id: "{id_projeto}", // código gerado automaticamente\n')
            arquivo_js.write('    passingScore: 80,\n')
            arquivo_js.write('    statusApproved: "completed",\n')
            arquivo_js.write('    statusDisapproved: "incomplete",\n')
            arquivo_js.write('    statusInit: "incomplete",\n')
            arquivo_js.write('    blockScreen: true,\n')
            arquivo_js.write('    bookMark: true,\n')
            arquivo_js.write('    navigation: "vertical",\n')
            arquivo_js.write('    menuVisible: false,\n\n')
            arquivo_js.write('    evaluation: {\n')
            arquivo_js.write('        showResultQuestion: true,\n')
            arquivo_js.write('        random: true,\n')
            arquivo_js.write('        randomOptions: true,\n')
            arquivo_js.write('        chances: false,\n')
            arquivo_js.write('        neverReAsk: false,\n')
            arquivo_js.write('        size: function () {\n')
            arquivo_js.write('            let sizeOfAssertment = 0;\n')
            arquivo_js.write('            let length = conf_mp.evaluation.assertment.length;\n')
            arquivo_js.write('            for (let i = 0; i < length; i++) {\n')
            arquivo_js.write('              sizeOfAssertment += conf_mp.evaluation.assertment[i].occur;\n')
            arquivo_js.write('            }\n')
            arquivo_js.write('            return sizeOfAssertment;\n')
            arquivo_js.write('          },\n')

            # Adicionar as categorias na lista "subjects"
            categorias = df['CATEGORIA/SEGMENTO (quando aplicável)'].dropna().unique()
            arquivo_js.write(f'        subjects: {list(categorias)},\n\n')

            # Estrutura da avaliação
            arquivo_js.write('        assertment: [\n')

            # Iterar pelas perguntas e adicionar ao JS
            for categoria in categorias:
                questoes_categoria = df[df['CATEGORIA/SEGMENTO (quando aplicável)'] == categoria]
                arquivo_js.write('            {\n')
                arquivo_js.write(f'                subject: "{categoria}",\n')
                arquivo_js.write(f'                occur: {len(questoes_categoria)},\n')
                arquivo_js.write('                exercises: [\n')

                for index, row in questoes_categoria.iterrows():
                    # Alternativas
                    alternativas = [
                        f"<p in-data='0'>{row['Alternativa A']}</p>",
                        f"<p in-data='0'>{row['Alternativa B']}</p>",
                        f"<p in-data='0'>{row['Alternativa C']}</p>",
                        f"<p in-data='0'>{row['Alternativa D']}</p>",
                    ]

                    # Marcar a alternativa correta
                    gabarito = row['GABARITO']
                    alternativas[ord(gabarito) - ord('A')] = alternativas[ord(gabarito) - ord('A')].replace("in-data='0'", "in-data='1'")

                    # Feedbacks
                    feedback_positivo = f"<h1>Muito bem! Você respondeu corretamente.</h1><p>{row['Feedback POSITIVO (quando aplicável)']}</p>"
                    feedback_negativo = f"<h1>Que pena! Você não respondeu corretamente.</h1><p>{row['Feedback NEGATIVO (quando aplicável)']}</p>"

                    # Adicionar exercício
                    arquivo_js.write('                    {\n')
                    arquivo_js.write('                        specie: "na",\n')
                    arquivo_js.write(f'                        statement: "{row["PERGUNTA"]}",\n')
                    arquivo_js.write('                        alternatives: [\n')
                    for alternativa in alternativas:
                        arquivo_js.write(f'                            "{alternativa}",\n')
                    arquivo_js.write('                        ],\n')
                    arquivo_js.write('                        feedbacks: [\n')
                    arquivo_js.write(f'                            "{feedback_positivo}",\n')
                    arquivo_js.write(f'                            "{feedback_negativo}"\n')
                    arquivo_js.write('                        ],\n')
                    arquivo_js.write('                    },\n')

                # Fechar exercícios
                arquivo_js.write('                ]\n')
                arquivo_js.write('            },\n')

            # Fechar assertment e avaliação
            arquivo_js.write('        ]\n')
            arquivo_js.write('    }\n')
            arquivo_js.write('};\n')

        messagebox.showinfo("Sucesso", f"Arquivo {caminho_js} criado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível gerar o arquivo JS: {e}")

def selecionar_arquivo(codigo_projeto):
    caminho_excel = filedialog.askopenfilename(
        title="Selecione a planilha Excel",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
    )
    if caminho_excel:
        # Gerar o ID do projeto a partir do código fornecido
        id_projeto = gerar_id_projeto(codigo_projeto)
        # Chamar a função que gera o JS
        excel_para_js(caminho_excel, 'conf.js', id_projeto)

# Criar a janela principal da interface gráfica
def iniciar_interface():
    janela = tk.Tk()
    janela.title("Gerador avaliação Micropower")
    janela.geometry("400x250")

    # Campo para inserir o código do projeto
    label_codigo = tk.Label(janela, text="Código do Projeto (ex: SIC028E-28):")
    label_codigo.pack(pady=10)
    entrada_codigo = tk.Entry(janela)
    entrada_codigo.pack(pady=5)

    # Botão para selecionar a planilha e gerar o JS
    def ao_clicar():
        codigo_projeto = entrada_codigo.get()
        if not codigo_projeto:
            messagebox.showerror("Erro", "Por favor, insira o código do projeto.")
        else:
            selecionar_arquivo(codigo_projeto)

    botao = tk.Button(janela, text="Selecionar Planilha", command=ao_clicar)
    botao.pack(pady=20)

    janela.mainloop()

# Iniciar a interface gráfica
if __name__ == "__main__":
    iniciar_interface()
