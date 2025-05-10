"""
Módulo de automação para planejamento de produção.
"""

# Expondo as funções principais de production_planner
from automation.core.production_planner import preencher_producao

# Expondo as funções de actions
from automation.actions.action_selector import escolher_acao
from automation.actions.add_row import adicionar_nova_linha
from automation.actions.remove_order import excluir_pedido
from automation.actions.create_plan import criar_novo_plano
from automation.actions.priority_handler import definir_prioridade, definir_ordem_manual
from automation.actions.reports_export import gerar_relatorio_arquivo

# Expondo funções de ui
from automation.ui.file_selector import escolher_arquivo_excel, escolher_arquivo_exportar
from automation.ui.cut_selector import selecionar_tipos_de_corte
from automation.ui.table_renderer import processar_tabela

# Expondo validadores
from automation.validators.report_validator import validar_prazo
from automation.validators.start_date_validator import validar_data_input

# Expondo constantes
from automation.core.constants import SETOR_ORDEM

__all__ = [
    'preencher_producao',
    'escolher_acao',
    'adicionar_nova_linha',
    'excluir_pedido',
    'criar_novo_plano',
    'definir_prioridade',
    'definir_ordem_manual',
    'gerar_relatorio_arquivo',
    'escolher_arquivo_excel',
    'escolher_arquivo_exportar',
    'selecionar_tipos_de_corte',
    'processar_tabela',
    'validar_prazo',
    'validar_data_input',
    'SETOR_ORDEM'
]