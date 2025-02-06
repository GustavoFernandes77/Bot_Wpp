# Sistema de Gerenciamento de Clientes e Envio de Mensagens Automáticas

Este projeto consiste em um sistema para gerenciar clientes e enviar mensagens automáticas via WhatsApp com base no status de pagamento e na data de vencimento de boletos. O sistema é composto por dois módulos principais: `main.py` (para gerenciamento de clientes) e `app.py` (para envio de mensagens).

---

## 📋 Visão Geral

O sistema é projetado para automatizar o envio de lembretes de pagamento para clientes, reduzindo o trabalho manual e garantindo que os clientes sejam notificados de forma consistente. Ele utiliza uma planilha Excel (`CLIENTES.xlsx`) para armazenar os dados dos clientes e uma interface gráfica para interação com o usuário.

---

## 🚀 Funcionalidades

### 1. Gerenciamento de Clientes (`main.py`)
- **Adicionar clientes**: Inserir novos clientes com informações como nome, telefone, CPF, e-mail, endereço, data de vencimento e status de pagamento.
- **Editar clientes**: Modificar informações de clientes existentes.
- **Remover clientes**: Excluir clientes da lista.
- **Visualizar clientes**: Exibir a lista de clientes em uma tabela.

### 2. Envio de Mensagens Automáticas (`app.py`)
- **Carregar clientes**: Carregar os dados dos clientes da planilha Excel.
- **Filtrar clientes**: Filtrar clientes que ainda não pagaram e têm uma data de vencimento.
- **Enviar mensagens**: Enviar mensagens personalizadas via WhatsApp Web.
- **Personalização**: Permitir a personalização das mensagens antes do envio.

---

## 📝 Requisitos

### Requisitos Funcionais
1. Gerenciamento de clientes (adicionar, editar, remover).
2. Envio de mensagens automáticas via WhatsApp.
3. Persistência de dados em uma planilha Excel (`CLIENTES.xlsx`).

### Requisitos Não Funcionais
1. **Usabilidade**: Interface gráfica intuitiva e responsiva.
2. **Desempenho**: Carregamento rápido e envio eficiente de mensagens.
3. **Segurança**: Validação de dados para evitar entradas inválidas.
4. **Confiabilidade**: Tratamento de erros robusto e registro de falhas.

---

## 🛠️ Melhorias Propostas

1. **Validação de Dados**:
   - Adicionar validações para telefone, CPF e e-mail.
   - Garantir que os dados estejam no formato correto antes de salvar.

2. **Automatização do Login no WhatsApp**:
   - Utilizar a biblioteca `selenium` para automatizar o login no WhatsApp Web.

3. **Melhor Tratamento de Erros**:
   - Implementar retentativas para envio de mensagens que falharam.
   - Adicionar logs detalhados para facilitar a depuração.

4. **Persistência em Banco de Dados**:
   - Migrar para um banco de dados SQLite ou MySQL para melhorar a escalabilidade.

5. **Interface Gráfica Mais Amigável**:
   - Adicionar feedback visual (indicadores de progresso).
   - Melhorar o layout e a usabilidade geral.

---

# 🖥️ Como Usar

## 📌 Pré-requisitos
- Python 3.x instalado.
- Bibliotecas necessárias: `openpyxl`, `pyautogui`, `customtkinter`, `tkinter`.

## 📌 Instalação das Bibliotecas
Execute o seguinte comando para instalar as dependências:

```bash
pip install openpyxl pyautogui customtkinter
```

---

# 🖥️ Executando o Sistema

## 📌 Gerenciamento de Clientes
Execute o arquivo `main.py` para gerenciar clientes.

```bash
python main.py
```
![main_captura](https://github.com/user-attachments/assets/43300cbd-5101-4ec6-8bca-867241a102a5)

## 📌 Envio de Mensagens
Execute o arquivo `app.py` para enviar mensagens automáticas.

```bash
python app.py
```
![app_captura](https://github.com/user-attachments/assets/23482b21-f6e9-4022-8fa0-f2afb7ca4a01)

---

## 📂 Estrutura do Projeto
```
projeto/
├── main.py                # Interface de gerenciamento de clientes
├── app.py                 # Interface de envio de mensagens
├── CLIENTES.xlsx          # Planilha de armazenamento de dados
├── erro.csv               # Arquivo de registro de erros (gerado automaticamente)
├── README.md              # Documentação do projeto
├── LICENSE                # Arquivo da licença MIT
```

---

## 📄 Licença
Este projeto é licenciado sob a licença MIT.

---

## 🤝 Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para abrir *issues* ou enviar *pull requests*.


