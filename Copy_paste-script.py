import time
import pyautogui
import pandas as pd
  
#1 etapa: coletar dados do csv e cadastrar auditorias    
try:
    #leitura do banco de dados csv e armazenamento
    audit_table = pd.read_csv("Auditorias.csv", dtype={"Rej. Int. (%)": str, "Lotes Reprovados (%)": str})
    
    #pode utilizar o print para mostrar tabela
    #print(products_table)

    #para coletar os dados do csv, utilizamos o loop em FOR
    for line in audit_table.index:
        #inicio de cadastro das auditorias
        #click campo de data
        time.sleep(0.5)
        pyautogui.click(x=1204, y=482)
        pyautogui.write(audit_table.loc[line, "Dia"])

        #campo data
        time.sleep(0.5)
        pyautogui.write(audit_table.loc[line, "Mês"])
        
        #campo data
        time.sleep(0.5)
        pyautogui.write(audit_table.loc[line, "Ano"])

        #campo rej int
        time.sleep(0.5)
        pyautogui.press('tab')
        rej_int = str(audit_table.loc[line, "Rej. Int. (%)"])
        pyautogui.write(rej_int)

        #campo lotes reprovados
        time.sleep(0.5)
        pyautogui.click(x=1238, y=688)
        lotes_reprovados = str(audit_table.loc[line, "Lotes Reprovados (%)"])
        pyautogui.write(lotes_reprovados)

        #registrar auditoria
        time.sleep(0.5)
        pyautogui.click(x=1042, y=769)
        pyautogui.press('enter')

    
except Exception as e:
    print(f"Ocorreu um erro: {e}")
    print("O programa foi encerrado")
