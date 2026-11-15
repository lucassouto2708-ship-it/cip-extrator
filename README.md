# Extrator CIP

Aplicativo em Streamlit que lê os PDFs de pagamento da Contribuição para o Custeio da Iluminação Pública (COSIP, código 117), corrige os valores pela Selic e exporta uma planilha Excel com a auditoria.

## Como usar

**Windows:** dê dois cliques em `rodar.bat`. Ele abre o app no navegador.

**Outros sistemas:**

```bash
pip install streamlit pandas pdfplumber openpyxl requests
streamlit run app.py
```

Depois:
1. Envie um ou mais PDFs de relatório de pagamentos.
2. Clique em **Processar e adicionar**.
3. Confira o resumo e os períodos processados.
4. Baixe a planilha Excel, com a correção pela Selic aplicada.

## Como funciona

- Cada PDF tem o mês e o ano identificados, e o valor da COSIP é localizado pelo código 117 ou pela rubrica 1.751.000.
- A Selic mensal vem da série 4390 do Banco Central, consultada na API do BCB. Se a API estiver fora, o código usa uma tabela de reserva embutida em `app.py` (`SELIC_FALLBACK`).

## Arquivos

- `app.py`: todo o aplicativo (leitura dos PDFs, correção pela Selic, interface e geração do Excel).
- `rodar.bat`: atalho para abrir o app no Windows.
- `.streamlit/config.toml`: tema visual e desligamento das estatísticas de uso do Streamlit.

## Atenção

- A tabela `SELIC_FALLBACK` precisa ser atualizada quando novos meses forem publicados, caso a API do BCB não responda.
- O app lê apenas a primeira página de cada PDF.
- Não há arquivo de dependências no repositório. A lista de pacotes acima foi tirada do código.
