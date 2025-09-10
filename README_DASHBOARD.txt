
# Dashboard de Vistorias (Excel → Python/Streamlit)

## Como rodar (Windows/macOS)
1) Abra um terminal na pasta onde baixou estes arquivos.
2) (Opcional) Crie um ambiente virtual:
   - Windows (PowerShell):
     python -m venv .venv
     .\.venv\Scripts\Activate.ps1
   - macOS/Linux:
     python3 -m venv .venv
     source .venv/bin/activate
3) Instale as dependências:
   pip install -r requirements_dashboard.txt
4) Execute o app:
   streamlit run app_dashboard_vistorias.py
   (Se der erro 'streamlit' não reconhecido, use: python -m streamlit run app_dashboard_vistorias.py)
5) No navegador, abra o link exibido (geralmente http://localhost:8501).

## Como usar
- O app tenta ler automaticamente o arquivo **Acomp. de Vistorias CRO1 - 2025.xlsx** se ele estiver na mesma pasta.
- Você também pode enviar qualquer .xlsx pelo upload do app.
- Há filtros por período (com base na Data da Solicitação), Diretoria, Situação, Urgência e OM.

## KPIs e Gráficos
- Total de vistorias, % finalizadas, prazo médio total e de execução.
- Evolução mensal, barras por Diretoria, pizza por Situação, barras por Urgência e tabela detalhada.
