import os
import win32com.client as win32

# Caminho da pasta onde o script está
pasta = os.path.dirname(os.path.abspath(__file__))

# Abre o PowerPoint
ppt = win32.Dispatch("PowerPoint.Application")
ppt.Visible = 1  # 1 = visível, 0 = oculto

for arquivo in os.listdir(pasta):
    if arquivo.lower().endswith(".pptx"):
        caminho_pptx = os.path.join(pasta, arquivo)
        nome_pdf = arquivo.replace(".pptx", ".pdf")
        caminho_pdf = os.path.join(pasta, nome_pdf)

        print(f"Convertendo: {arquivo}")

        # Abre a apresentação
        apresentacao = ppt.Presentations.Open(caminho_pptx, WithWindow=False)

        # Exporta para PDF
        apresentacao.SaveAs(caminho_pdf, FileFormat=32)  # 32 = PDF
        apresentacao.Close()

ppt.Quit()

print("Conversão concluída!")
