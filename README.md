# Robô de Monitoramento Diário de Preço de um Produto

Este script Python automatiza o monitoramento diário do preço de um produto 
específico em um site, gravando os dados em uma planilha Excel.

## Funcionalidades

* Acessa um site definido (url) e extrai o preço de um produto.
* Processa o preço extraído, convertendo-o para inteiro ou decimal.
* Cria uma planilha Excel com informações sobre o produto, data e hora da 
  coleta, preço e link do produto.
* Insere os dados coletados na planilha Excel.
* Agenda a execução do script em intervalos regulares (definido em minutos).

## Requisitos

* Python 3.12.0
* Browser Google Chrome instalado

## Como usar

1. **Clone o repositório:**
   ```bash
   git clone https://github.com/eliasalbuquerque/python-price-tracker-bot
   ```

2. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Executar o script:**
   ```bash
   python app.py
   ```
   * O script irá iniciar a coleta de dados e salvará os dados na planilha.
   * O script será agendado para execução a cada 30 minutos padrão.
   * Para interromper o script, pressione a tecla `ESC` no teclado.

## Notas

* Este script é um exemplo básico e pode ser personalizado para atender 
  necessidades específicas.
* O script utiliza o `keyboard` para monitorar a tecla `ESC` e interromper a 
  execução, o que pode não funcionar em todos os ambientes.
* A função `access_website()` espera 15 segundos para que a página carregue 
  completamente. Ajuste este tempo caso seja necessário.
* O script possui um determinado `XPATH` para o seu funcionamento. Ajuste a 
  função `extract_product_value()` do script caso o site seja diferente.

## Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues, enviar pull 
requests ou sugerir melhorias.

## Funcionamento do script

https://youtu.be/XVwewpsmejI
