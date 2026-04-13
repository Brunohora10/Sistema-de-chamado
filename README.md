# Sistema de Chamados CadÚnico

Aplicação web em Google Apps Script para gestão de senhas e atendimento do CadÚnico, com telas de recepção, guichê, painel TV e administração.

## Visão Geral

Este sistema foi construído para organizar o fluxo de atendimento em setores com alta demanda, garantindo:

- emissão rápida de senhas por tipo de atendimento;
- chamada ordenada por prioridade configurável;
- operação de múltiplos guichês e setores;
- exibição em painel TV com anúncio por voz;
- administração de parâmetros do sistema em tempo real;
- histórico e estatísticas para gestão.

## Principais Funcionalidades

- Recepção
  - cadastro de senha com nome, CPF, bairro e serviço;
  - impressão de ticket;
  - lista de senhas aguardando.
- Guichê
  - chamar próximo;
  - repetir chamada;
  - iniciar/finalizar atendimento;
  - marcar não compareceu;
  - chamada manual.
- Painel TV
  - destaque da última chamada;
  - grade com últimas senhas chamadas;
  - áudio TTS com repetição configurável;
  - retomada de estado após fechar/reabrir a aba.
- Admin
  - configurações de áudio, guichês e impressão;
  - serviços e bairros configuráveis;
  - ordem de chamada (N, P, G);
  - PIN da Tela TV;
  - monitoramento ao vivo e estatísticas.

## Estrutura do Projeto

- `code.js`: backend Apps Script (regras, persistência e APIs)
- `index.html`: frontend (UI, estados e integrações)
- `appsscript.json`: manifesto do Apps Script
- `.clasp.json`: vínculo local com o projeto Apps Script

## Requisitos

- Node.js 18+ (recomendado)
- npm
- `@google/clasp` instalado globalmente
- acesso ao projeto do Google Apps Script
- acesso à planilha vinculada no `SPREADSHEET_ID`

## Setup Local

### 1. Instalar clasp

```bash
npm i -g @google/clasp
```

### 2. Login no Google

```bash
clasp login
```

### 3. Clonar projeto Apps Script (se necessário)

```bash
clasp clone <SCRIPT_ID>
```

### 4. Enviar alterações locais

```bash
clasp push
```

### 5. Abrir editor Apps Script

```bash
clasp open
```

## Deploy Web App

No editor do Apps Script:

1. Deploy > New deployment
2. Tipo: Web app
3. Execute as: User accessing the web app
4. Who has access: Anyone (ou conforme política da instituição)
5. Deploy e copie a URL

## Configuração Inicial Recomendada

1. Acessar tela Admin
2. Revisar:
   - quantidade de guichês;
   - tempo de chamada;
   - volume/velocidade/repetições de áudio;
   - lista de serviços;
   - lista de bairros;
   - ordem de chamada.
3. Definir PIN da Tela TV
4. Validar fluxo completo:
   - gerar senha;
   - chamar no guichê;
   - conferir exibição na TV;
   - iniciar/finalizar atendimento.

## Regras de Fluxo Importantes

- Um guichê não pode iniciar nova chamada enquanto houver chamada pendente ou atendimento em andamento no mesmo guichê.
- Timeouts automáticos encerram estados travados (atendendo/chamando) após limite configurado no backend.
- Chamada manual respeita setor e status da senha.

## Performance

O projeto já inclui otimizações para reduzir lentidão:

- cache curto no backend para dados do dia;
- invalidação de cache em operações de escrita;
- polling sem sobreposição no frontend;
- atualização de UI apenas quando os dados mudam;
- snapshot local para acelerar primeira renderização.

## Troubleshooting

### A TV não anuncia chamada

- confirmar que a aba da TV está aberta na tela correta;
- clicar em "Ativar som" no overlay;
- validar permissão/autoplay do navegador;
- confirmar se a chamada entrou como status `chamando` no backend.

### Lentidão ao carregar

- verificar tamanho da planilha e histórico acumulado;
- validar se o deploy publicado é o mais recente;
- confirmar que as otimizações estão na versão ativa;
- monitorar erros no Console do navegador.

### Erro de acesso à planilha

- revisar `SPREADSHEET_ID` em `code.js`;
- confirmar permissão do usuário no Google Sheets;
- garantir deploy com execução correta para o contexto da equipe.

## Segurança e Boas Práticas

- não versionar segredos;
- usar PIN da TV diferente do padrão;
- revisar permissões de compartilhamento do Apps Script e da planilha;
- manter histórico de mudanças no GitHub com commits descritivos.

## Versionamento

Fluxo sugerido para alterações:

```bash
git add .
git commit -m "feat: descricao objetiva da mudança"
git push origin main
clasp push
```

## Licença

Definir conforme política da organização (ex.: uso interno institucional).
