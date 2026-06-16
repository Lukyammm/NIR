# Analise da planilha e dos rascunhos

Data da analise: 16/06/2026.

## Material analisado

- Workbook `CÓPIA DE TESTE APP NIR.xlsx`, com 44 abas.
- Rascunhos manuscritos sobre UIB, filtros, tempos, relatório de passagem de plantão, fast locais, permanência e desfechos.
- Histórico anterior do repositório, que indicava uma primeira versão em Google Apps Script com abas vivas, histórico e relatório editável.

## Estrutura encontrada na planilha

### Cadastros e apoio

- `RAMAIS E AVISOS`: ramais, setores, coordenações, staffs, contatos e avisos operacionais.
- Links úteis nas abas de plantão: Fastmedic, controle de altas da UTI, isolamento, perfil/pedidos, filas externas.

### Filas, procedimentos e rede externa

- `TRANSFERÊNCIA EXTERNA + HD + CC`: transferências externas, hemodiálise, casa de cuidados, SAD/EMAD/CCC, status de pendência, contatos e tempo aguardando.
- `ENDOSCOPIA (MED)`: endoscopia e procedimentos relacionados, com aptidão médica, avaliador, agendamento, status e observação.
- `VASCULAR + RADIO + URO (MED)` e `RESOLVIDOS...`: agenda de procedimentos externos, judicialização, aptidão, avaliador, status e observação.
- `IMAGEM (ADM + MED)`: exames de imagem, agendamento, avaliador, status, observações de sedação/reagendamento.
- `FAST LOCAIS (ADM+ENF)`: solicitações locais/ambulatoriais, clínica, leito, prontuário, origem e situação.

### Ocupação, UIB e bloqueios

- `UIBS (ENF + MED)` e `imp`: ocupação por unidade, leito, nome/prontuário/fast, origem, especialidade, observação, pendências, permanência e status como livre, reservado, falta chegar e internado.
- `LEITOS ECTÓPICOS + BLOQUEADO`: leitos ectópicos, bloqueios por manutenção, isolamento e pacientes fora da unidade ideal.

### Plantões diários

As abas por data/turno (`01062026 MT`, `01062026 SN`, ..., `16062026 MT`) têm um padrão:

- Cabeçalho do plantão: data, turno, dia da semana, vagas MT/SN, equipe médica, enfermagem, administrativo e coordenação.
- Quadro de vagas por especialidade: internações eletivas programadas, altas programadas do plantão SN anterior, manhã/tarde/noite.
- Fila de internação Dona Libânia: origem da regulação, Fastmedic, nome, data, especialidade, unidade de origem, status, avaliador, justificativa, observação e sincronização.
- Plantão anterior a acompanhar: registros ainda não concluídos.
- Pedidos de UTI, SRPA aguardando leito, altas da UTI aguardando leito e altas com leito.
- Pendências livres e campos de assinatura/carimbo.

## O que os rascunhos acrescentam

Os rascunhos pedem que o app não seja apenas uma cópia da planilha. Pontos essenciais:

- Filtros por período, turno (`MT`, `SN` e multiseleção), status, especialidade, ADS/município, unidade de origem, tipo de regulação e UIB.
- Relatório de passagem de plantão como função principal.
- Controle de tempo em marcos: vaga dada, paciente reservado, aceite, chegada, internação/cancelamento.
- UIB com tempo de permanência, média de permanência, desfecho, tipos de desfecho e possibilidade de transferência entre UIBs.
- Sinalização de paciente em UIB errada conforme especialidade.
- Fast locais como aba/módulo próprio.
- Gráficos por dia/mês e evolução de vagas/status.
- Diferenciar informações momentâneas daquelas que precisam de histórico e auditoria.

## Problemas da planilha atual

- Dados de naturezas diferentes ficam misturados na mesma aba.
- Muitas abas por data/turno dificultam análise longitudinal e aumentam retrabalho.
- Status são escritos de várias formas, o que prejudica filtro e indicador.
- O mesmo paciente/evento pode aparecer em fila, UIB, UTI e procedimento sem vínculo único.
- Tempos importantes não são sempre carimbados no momento exato.
- Desfecho e permanência em UIB não estão padronizados.
- Relatório de plantão é parcialmente manual e depende de copiar contexto de várias regiões da planilha.
- Há risco de privacidade se planilhas ou exports forem compartilhados fora do ambiente privado.

## Decisões propostas

- Manter Google Sheets como banco do MVP, mas com abas normalizadas.
- Parar de criar aba nova por plantão como fonte primária; gerar o relatório por filtro a partir das tabelas.
- Cada registro deve ter `id`, `plantao_id`, `created_at`, `updated_at`, status padronizado e campos de auditoria.
- Dados reais ficam somente em planilha privada. O repositório usa exemplos sintéticos.
- Relatório diário deve ser gerado por filtros e salvo como snapshot ao encerrar plantão.
