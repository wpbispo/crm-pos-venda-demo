# CRM de pós-venda — loja de purificadores (demonstração)

CRM que roda direto no navegador, sem instalação e sem servidor, para organizar o contato com clientes que precisam trocar o refil do purificador de água.

> **Dados fictícios.** Esta é uma versão de demonstração. A loja "Fonte Clara", os clientes, telefones e endereços foram inventados. Os telefones usam o DDD 00, que não existe.

**Testar online:** https://wpbispo.github.io/crm-pos-venda-demo/

## O problema

Uma loja de purificadores vende refis que vencem a cada 6 ou 12 meses. O controle de quem precisava trocar ficava em abas mensais de uma planilha Excel. Era difícil saber quem já tinha sido contatado, o que o cliente respondeu e quando voltar a falar com ele.

## O que o CRM faz

- **Fila de hoje:** mostra quem está com o refil vencido ou vencendo, em colunas (para contatar, em contato, agendado).
- **Registro de cada contato:** situação da conversa, anotação ("vai esperar o cartão virar") e data para voltar a falar.
- **Agendamentos:** calendário para marcar trocas e instalações, arrastando o cliente para o dia.
- **WhatsApp:** abre a conversa com uma mensagem pronta, de acordo com a situação do cliente.
- **Mutirão:** passa por vários clientes em sequência, pulando quem já recebeu mensagem nos últimos dias.
- **Lista para o celular:** gera uma página para fazer as ligações pelo celular e depois trazer os resultados de volta.
- **Atualizar clientes:** lê a planilha Excel da loja direto no navegador. Use a `Planilha de Clientes (exemplo).xlsx` desta pasta para testar.
- **Backup:** salva e restaura tudo em um arquivo `.json`.
- Tema claro e escuro e layout para celular.

## Tecnologias

- HTML, CSS e JavaScript puros, sem framework.
- [SheetJS](https://sheetjs.com/) para ler arquivos `.xlsx` no navegador.
- `localStorage` do navegador para guardar as situações, anotações e agendamentos.
- Git e GitHub Pages para publicar.

## Como rodar

Baixe a pasta e abra o `index.html` no Chrome. Não precisa instalar nada.

## Como foi feito

Desenvolvido com apoio de IA (Claude). Eu levantei as necessidades a partir da rotina de pós-venda que eu mesmo faço na loja, defini como cada tela deveria funcionar, testei e uso a versão real no dia a dia. A versão real lê a planilha verdadeira da loja e fica só no computador da empresa.

---
Wiliam Bispo · [LinkedIn](https://www.linkedin.com/in/wiliambispo/)
