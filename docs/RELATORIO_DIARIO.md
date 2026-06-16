# Relatório diário de troca de plantão

## Objetivo

Gerar, quando necessário, um relatório diário de passagem de plantão do NIR com filtros por período, turno, origem, especialidade, status, unidade, ADS/município e busca livre.

O relatório deve reduzir digitação manual e manter rastreabilidade do que estava pendente no momento da passagem.

## Filtros mínimos

- Data inicial e final
- Turno: `MT`, `SN` ou todos
- Módulos: regulação, UIB/leitos, procedimentos, UTI/SRPA, bloqueios
- Status
- Especialidade
- Origem da regulação: CRL, SAMU, local, SIREG, ambulatório
- ADS/regional
- Município
- Unidade de origem
- Profissional avaliador
- Unidade/leito
- Somente pendentes
- Busca livre

## Estrutura do relatório

1. Identificação do plantão
   - data
   - turno
   - equipe médica
   - equipe de enfermagem
   - administrativos
   - coordenação

2. Resumo executivo
   - total de regulações
   - reservados
   - internados
   - negados
   - cancelados
   - pendências abertas
   - procedimentos pendentes
   - leitos bloqueados
   - altas de UTI aguardando leito

3. Vagas e especialidades
   - vagas por especialidade
   - altas programadas
   - internações eletivas programadas
   - vagas remanescentes

4. Regulações do plantão
   - origem
   - Fastmedic
   - identificação segura
   - especialidade
   - unidade de origem
   - status
   - avaliador
   - justificativa
   - observação

5. Plantão anterior a acompanhar
   - apenas registros não concluídos
   - prioridade
   - motivo de permanência
   - próximo passo

6. UIB / Leitos
   - unidade
   - leito
   - status
   - paciente
   - especialidade de origem e destino
   - permanência
   - pendência
   - desfecho
   - alerta de UIB potencialmente inadequada

7. Procedimentos e rede externa
   - imagem
   - endoscopia
   - vascular/radio/urologia
   - hemodiálise
   - transferência externa
   - SAD/EMAD/CCC

8. UTI / SRPA
   - pedidos de UTI
   - SRPA aguardando leito
   - altas da UTI aguardando leito
   - altas da UTI com leito

9. Bloqueios e isolamento
   - manutenção
   - isolamento
   - ectópicos
   - previsão de liberação

10. Pendências livres
    - assuntos que não cabem nas tabelas, mas precisam ser passados.

## Regras de geração

- O relatório deve mostrar primeiro o que exige ação.
- Status finais podem aparecer em resumo, mas registros concluídos devem ser omitidos quando o filtro "somente pendentes" estiver ativo.
- Ao encerrar plantão, salvar snapshot do texto e dos dados-base.
- Permitir copiar texto para WhatsApp/e-mail e imprimir em PDF pelo navegador.
- Manter o relatório reproduzível: o usuário deve conseguir saber quais filtros geraram aquele texto.

## Métricas de tempo

- `vaga_dada_em -> reservado_em`
- `reservado_em -> chegou_em`
- `chegou_em -> internado_em`
- `reservado_em -> cancelado_em`
- permanência em UIB

Esses tempos devem ser calculados automaticamente quando os marcos forem preenchidos.
