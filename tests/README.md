# Validação da importação de programação

Com Node.js, Playwright e Chromium disponíveis, inicie `python3 -m http.server 8000 --bind 127.0.0.1` na raiz do repositório e execute `node tests/schedule-import.cjs` em outro terminal. O teste usa Chromium em `/usr/bin/chromium` e contextos descartáveis; não altera os dados do navegador do usuário.

Os arquivos Excel e CSV são gerados em memória. A validação cobre as duas páginas, seleção de aba, datas, aliases de região, equipes novas, dados obrigatórios ausentes, repetição no arquivo e entre importações, preservação de registros manuais e rechecagem ao confirmar.

O leitor local ExcelJS 4.4.0 vem do pacote npm oficial `exceljs@4.4.0`, instalado com verificação de integridade pelo npm. Sua licença MIT está em `assets/vendor/exceljs-LICENSE`. Os arquivos são processados no navegador; não há envio da planilha para serviço de conversão. A sincronização de agendamentos existente continua seguindo a configuração do aplicativo.
