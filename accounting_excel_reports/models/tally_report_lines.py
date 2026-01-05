from odoo import models, fields

class TallyTrialBalanceLine(models.TransientModel):
    _name = 'tally.trial.balance.line'
    _description = 'Tally Style Trial Balance Line (Transient)'
    _order = 'sequence, id'

    wizard_id = fields.Many2one('trial.balance.wizard', ondelete='cascade')
    sequence = fields.Integer(default=10)
    level = fields.Integer(string='Level', default=0)
    name = fields.Char(string='Particulars')
    debit = fields.Float(string='Debit', digits='Account')
    credit = fields.Float(string='Credit', digits='Account')
    is_group = fields.Boolean(string='Is Group')
    is_total = fields.Boolean(string='Is Total')
    
    # --- Drill-down Fields ---
    account_id = fields.Many2one('account.account', string='Account')  # For fetching journal details
    is_expandable = fields.Boolean(string='Has Expandable Details', default=False)
    is_child_line = fields.Boolean(string='Is Child Detail Line', default=False)
    is_detail_line = fields.Boolean(string='Is Journal Detail Line', default=False)
    group_key = fields.Char(string='Group Key')
    parent_key = fields.Char(string='Parent Key')  # Links detail line to its parent account
    # --- End Drill-down Fields ---

class TallyProfitLossLine(models.TransientModel):
    _name = 'tally.profit.loss.line'
    _description = 'Tally Style Profit & Loss Line (Transient)'
    _order = 'sequence, id'

    # Vertical Link
    wizard_id = fields.Many2one('profit.loss.wizard', ondelete='cascade')
    
    # --- New Fields for Horizontal Report ---
    wizard_income_id = fields.Many2one('profit.loss.wizard', ondelete='cascade', string="Wizard (Income)")
    wizard_expense_id = fields.Many2one('profit.loss.wizard', ondelete='cascade', string="Wizard (Expenses)")
    # ----------------------------------------

    sequence = fields.Integer(default=10)
    level = fields.Integer(string='Level', default=0)
    name = fields.Char(string='Particulars')
    code = fields.Char(string='Code')
    amount = fields.Float(string='Amount', digits='Account')
    is_group = fields.Boolean(string='Is Group')
    is_total = fields.Boolean(string='Is Total')
    is_net_result = fields.Boolean(string='Is Net Profit/Loss')
    
    # --- Drill-down Fields ---
    account_id = fields.Many2one('account.account', string='Account')
    is_expandable = fields.Boolean(string='Has Expandable Details', default=False)
    is_child_line = fields.Boolean(string='Is Child Detail Line', default=False)
    is_detail_line = fields.Boolean(string='Is Journal Detail Line', default=False)
    group_key = fields.Char(string='Group Key')
    parent_key = fields.Char(string='Parent Key')
    # --- End Drill-down Fields ---

class TallyBalanceSheetLine(models.TransientModel):
    _name = 'tally.balance.sheet.line'
    _description = 'Tally Style Balance Sheet Line (Transient)'
    _order = 'sequence, id'

    # For Vertical Report
    wizard_id = fields.Many2one('balance.sheet.wizard', ondelete='cascade')
    
    # --- New Fields for Horizontal Report ---
    wizard_liab_id = fields.Many2one('balance.sheet.wizard', ondelete='cascade', string="Wizard (Liabilities)")
    wizard_asset_id = fields.Many2one('balance.sheet.wizard', ondelete='cascade', string="Wizard (Assets)")
    # --- End New Fields ---
    
    # --- Drill-down Fields ---
    account_id = fields.Many2one('account.account', string='Account')
    is_expandable = fields.Boolean(string='Has Expandable Details', default=False)
    is_child_line = fields.Boolean(string='Is Child Detail Line', default=False)
    is_detail_line = fields.Boolean(string='Is Journal Detail Line', default=False)
    group_key = fields.Char(string='Group Key')
    parent_key = fields.Char(string='Parent Key')
    # --- End Drill-down Fields ---

    sequence = fields.Integer(default=10)
    level = fields.Integer(string='Level', default=0)
    name = fields.Char(string='Particulars')
    code = fields.Char(string='Code')
    amount = fields.Float(string='Amount', digits='Account')
    is_group = fields.Boolean(string='Is Group')
    is_total = fields.Boolean(string='Is Total')