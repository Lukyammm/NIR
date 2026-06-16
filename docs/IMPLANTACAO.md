# Implantação do MVP

## Ambiente recomendado

- Google Sheets privado como base de dados inicial.
- Google Apps Script vinculado à planilha ou standalone com `PLANILHA_ID`.
- Acesso restrito aos usuários do NIR.
- Repositório público somente com código e documentação, sem dados reais.

## Passo a passo

1. Criar uma planilha Google privada.
2. Abrir Apps Script.
3. Criar dois arquivos:
   - `Code.gs`
   - `index.html`
4. Copiar o conteúdo do repositório.
5. Em **Configurações do projeto > Propriedades do script**, criar:
   - `PLANILHA_ID`: ID da planilha privada.
6. Executar `criarEstruturaNIR()`.
7. Executar `seedDemoDataNIR()` apenas se quiser dados sintéticos para treinamento.
8. Publicar como Web App.

## Rotina operacional

1. Abrir plantão e preencher equipe.
2. Registrar regulações, leitos, procedimentos, UTI/SRPA e bloqueios durante o turno.
3. Usar filtros para revisar pendências.
4. Gerar relatório de passagem.
5. Copiar/imprimir relatório conforme necessidade.
6. Encerrar plantão para salvar snapshot.

## Permissões

Perfis sugeridos:

- Administrador: cria estrutura, ajusta cadastros, gerencia acesso.
- Médico regulador: avalia, altera status, registra justificativas e observações clínicas.
- Enfermagem NIR: atualiza chegada, leito, UIB, permanência, desfecho e pendências.
- Administrativo NIR: atualiza Fastmedic, agendamento, contatos, documentos e sincronizações.
- Coordenação: consulta dashboards, relatórios e auditoria.

## Limitações do MVP

- Google Sheets não substitui banco transacional em alto volume.
- Controle fino de autenticação depende do ambiente Google Workspace.
- Integração automática com Fastmedic/SAMU/SIREG não está implementada nesta fase.
- O app pressupõe digitação/atualização manual ou importação futura.

## Próximas fases

1. Importador da planilha histórica para as abas normalizadas.
2. Validações fortes por campo e listas institucionais.
3. Dashboard longitudinal com gráficos de dia, mês, especialidade, ADS e município.
4. Alertas de UIB incompatível com especialidade.
5. Integração com autenticação institucional e banco relacional.
6. Exportação PDF server-side com assinatura/versão do relatório.
