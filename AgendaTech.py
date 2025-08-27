import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox
import os


def gerar_agenda():
    try:
        # === Selecionar arquivo origem ===
        arquivo_origem = filedialog.askopenfilename(
            title="Selecione a planilha de origem",
            filetypes=[("Excel files", "*.xls;*.xlsx")]
        )
        if not arquivo_origem:
            return

        arquivo_saida = os.path.join(os.path.dirname(arquivo_origem), "AGENDA_FILTRADA.xlsx")

        # === Ler a planilha de origem ===
        df_origem = pd.read_excel(arquivo_origem)

        # === Lista de técnicos desejados ===
        nome_tecnico = ["ANDRE", "ADAUTO", "JOSIEL", "GUEDES", "DIOGO", "NATALICIO",
                        "ROMERO", "ESDRAS", "BRITO", "JUNIOR", "CILAS", "RONALDO", "MILTON", "JOAO", "SIDRAYTONN"]

        # === Filtrar os técnicos ===
        df_filtrado = df_origem[df_origem['Técnico'].isin(nome_tecnico)]

        # === Lista para armazenar os dados formatados ===
        linhas_formatadas = []

        for tecnico in nome_tecnico:
            df_tecnico = df_filtrado[df_filtrado['Técnico'] == tecnico].copy()

            if not df_tecnico.empty:
                df_tecnico = df_tecnico.reset_index(drop=True)
                df_tecnico['Sequência'] = df_tecnico.index + 1

                # Linha de título do técnico
                linhas_formatadas.append({
                    'Sequência': '',
                    'Cliente': f'TÉCNICO: {tecnico}',
                    'Bairro': '',
                    'Número da O.S': '',
                    'Tipo': '',
                    'Técnico': ''
                })

                # Cabeçalho
                linhas_formatadas.append({
                    'Sequência': 'Posição',
                    'Cliente': 'Cliente',
                    'Bairro': 'Cidade/Bairro',
                    'Número da O.S': 'Número da O.S',
                    'Tipo': 'Status',
                    'Técnico': 'Técnico'
                })

                for _, row in df_tecnico.iterrows():
                    linhas_formatadas.append({
                        'Sequência': row['Sequência'],
                        'Cliente': row['Cliente (Razão)'][:35],
                        'Bairro': (row['Cidade'] + ' - ' + row['Bairro'])[:50],
                        'Número da O.S': row['Seq. O.S.'],
                        'Tipo': row['Tipo de Status'][:10],
                        'Técnico': tecnico
                    })

        # Criar DataFrame final
        df_final = pd.DataFrame(linhas_formatadas)

        # === Salvar Excel formatado ===
        with pd.ExcelWriter(arquivo_saida, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Agenda', startrow=2, header=False)

            workbook = writer.book
            worksheet = writer.sheets['Agenda']

            # Inserir logo
            worksheet.insert_image('A1', 'LOGO SOLIVETTI.jpg', {
                'x_offset': 20,
                'y_offset': 5,
                'x_scale': 0.15,
                'y_scale': 0.15
            })


            # === Formatos ===
            formato_cabecalho = workbook.add_format({
                'bold': True,
                'bg_color': '#4F81BD',
                'font_color': 'white',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })
            formato_tecnico = workbook.add_format({
                'bold': True,
                'bg_color': '#C6EFCE',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })
            formato_dados = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })
            formato_titulo = workbook.add_format({
                'bold': 1,
                'font_size': 16,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })
            formato_data = workbook.add_format({
                'bold': 1,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': "#FFFB05",
                'num_format': 'dd/mm/yyyy',
                'text_wrap': True,
                'border': 1,
                'font_color': 'red'
            })
            formato_esquerda = workbook.add_format({
                'bold': True,
                'valign': 'vcenter',
                'bg_color': "#FCC7F3",
                'border': 1
            })

            # Inserir título
            worksheet.merge_range('B1:D2', 'AGENDA', formato_titulo)
            worksheet.merge_range('A3:D3', 'Técnicos internos:', formato_esquerda)

            data_amanha = datetime.today() + timedelta(days=1)
            worksheet.merge_range('E1:E2', data_amanha.strftime('%d/%m/%Y'), formato_data)
            worksheet.write('E3', data_amanha.strftime('%A'), formato_data)

            # Cabeçalho
            for idx, coluna in enumerate(df_final.columns):
                col_letter = chr(65 + idx)
                worksheet.write(f'{col_letter}4', coluna, formato_cabecalho)

            linha_excel = 5
            for _, row in df_final.iterrows():
                if str(row['Cliente']).startswith('TÉCNICO:'):
                    worksheet.merge_range(
                        f'A{linha_excel}:F{linha_excel}',
                        row['Cliente'],
                        formato_tecnico
                    )
                else:
                    worksheet.write(f'A{linha_excel}', row['Sequência'], formato_dados)
                    worksheet.write(f'B{linha_excel}', row['Cliente'], formato_dados)
                    worksheet.write(f'C{linha_excel}', row['Bairro'], formato_dados)
                    worksheet.write(f'D{linha_excel}', row['Número da O.S'], formato_dados)
                    worksheet.write(f'E{linha_excel}', row['Tipo'], formato_dados)
                    worksheet.write(f'F{linha_excel}', row['Técnico'], formato_dados)
                linha_excel += 1

            for idx, coluna in enumerate(df_final.columns):
                col_letter = chr(65 + idx)
                max_len = max(df_final[coluna].astype(str).map(len).max(), len(str(coluna))) + 5
                worksheet.set_column(f'{col_letter}:{col_letter}', max_len)

        messagebox.showinfo("Sucesso", f"Arquivo gerado: {arquivo_saida}")
    except Exception as e:
        messagebox.showerror("Erro", str(e))


# === Criar Janela ===
janela = tk.Tk()
janela.title("Gerador de Agenda")
janela.geometry("300x150")

btn = tk.Button(janela, text="Selecionar Planilha e Gerar", command=gerar_agenda, height=2, width=30, bg="#4F81BD",
                fg="white")
btn.pack(pady=40)

janela.mainloop()
