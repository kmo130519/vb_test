Option Explicitnn'#########################################################n' Class: clsTimeStepn' DESC : Contains time steps. Basically time steps are days preset by the user.n'        But, if any the fixing date or dividend date is included in a period,n'        it will be seperated by the initializer.n'#########################################################nnPrivate date_() As DatenPrivate dt_() As DoublenPrivate is_fixing_date_() As BooleannPrivate is_first_fixing_date_() As BooleannnPrivate is_dividend_date_() As BooleannnPrivate time_step_ As IntegernnPublic no_of_steps As Integernnn'------------------------------------------------n' Initializen'------------------------------------------------nPublic Sub initialize(deal_ticket As clsCliquetDealTicket, div_schedule As clsDividendSchedule)nn    Dim inx As Integern    Dim first_fixing_checked As Booleann    Dim fixing_inx As Integern    n    Dim div_inx As Integern    nOn Error GoTo ErrorHandlern    n    ReDim date_(0 To 0) As Daten    ReDim inx_(0 To 0) As Integern    ReDim dt_(0 To 0) As Doublen    ReDim is_fixing_date_(0 To 0) As Booleann    ReDim is_dividend_date_(0 To 0) As Booleann    ReDim is_first_fixing_date_(0 To 0) As Booleannn        n    inx = 0n    first_fixing_checked = Falsen    fixing_inx = 0n    n    time_step_ = deal_ticket.time_stepn    n    date_(inx) = deal_ticket.current_daten    n    fixing_inx = deal_ticket.fixing_indexnn    If fixing_inx > 0 Thenn    n        first_fixing_checked = Truen        n    End Ifn        n    div_inx = div_schedule.ini_div_inxn    n    Do While date_(inx) < deal_ticket.maturity_daten        n        fixing_check inx, fixing_inx, first_fixing_checked, deal_ticketn        div_check inx, div_inx, div_schedulen        n        make_next_step deal_ticket, fixing_inx, div_schedule, div_inx, inxn        n    Loopn    n    '---------------------------------------------------n    'check if the last date is fixing daten    fixing_check inx, fixing_inx, first_fixing_checked, deal_ticketn    n    no_of_steps = UBound(date_)n    n    Exit Subn    nErrorHandler:nn    Err.Raise vbObjectError + 1002, "clsTimeStep.initialize :" & Chr(13) & Err.source, Err.description '1002 <-- initialization errorn    Exit Subn        nnEnd SubnnnnPublic Function get_date(inx As Integer) As Daten    nOn Error GoTo ErrorHandlern        n    get_date = date_(inx)n    n    Exit Functionn    nErrorHandler:nn    Err.Raise vbObjectError + 1002, "clsTimeStep.get_date :" & Chr(13) & Err.source, Err.description '1002 <-- initialization errorn    Exit Functionn        nEnd FunctionnnPublic Function dt(inx As Integer) As Doublen    nOn Error GoTo ErrorHandlern        n    dt = dt_(inx)n    n    Exit Functionn    nErrorHandler:nn    Err.Raise vbObjectError + 1002, "clsTimeStep.dt :" & Chr(13) & Err.source, Err.description '1002 <-- initialization errorn    Exit Functionn    nEnd FunctionnPublic Function is_first_fixing_date(inx As Integer) As Booleann    nOn Error GoTo ErrorHandlern        n    is_first_fixing_date = is_first_fixing_date_(inx)nn    n    Exit Functionn    nErrorHandler:nn    Err.Raise vbObjectError + 1002, "clsTimeStep.is_first_fixing_date :" & Chr(13) & Err.source, Err.description '1002 <-- initialization errorn    Exit Functionn    nEnd FunctionnnPublic Function is_fixing_date(inx As Integer) As Booleann    nOn Error GoTo ErrorHandlern    n    is_fixing_date = is_fixing_date_(inx)n    n    Exit Functionn    nErrorHandler:nn    Err.Raise vbObjectError + 1002, "clsTimeStep.is_fixing_date :" & Chr(13) & Err.source, Err.description '1002 <-- initialization errorn    Exit Functionn    nEnd FunctionnnPublic Function is_dividend_date(inx As Integer) As Booleann    nOn Error GoTo ErrorHandlern    n    is_dividend_date = is_dividend_date_(inx)n    n    Exit Functionn    nErrorHandler:nn    Err.Raise vbObjectError + 1002, "clsTimeStep.is_dividend_date :" & Chr(13) & Err.source, Err.description '1002 <-- initialization errorn    Exit Functionn    nEnd Functionnn'------------------------------------------------n' called by Initializen' calculate this step's values and create next stepn'------------------------------------------------nPrivate Sub make_next_step(deal_ticket As clsCliquetDealTicket, fixing_inx As Integer, div_schedule As clsDividendSchedule, div_inx As Integer, inx As Integer)nnn    Dim days_per_year As Integern    nOn Error GoTo ErrorHandlern    n    '-------------------------------n    'subject to improven    '-------------------------------n    days_per_year = DateAdd("Yyyy", 1, deal_ticket.current_date) - deal_ticket.current_daten    n        n    If deal_ticket.get_max_fixing_index < fixing_inx Thenn    n        time_step_ = deal_ticket.maturity_date - date_(inx)n        dt_(inx) = time_step_ / days_per_yearn        n    Elsen    n        If date_(inx) + deal_ticket.time_step > deal_ticket.get_fixing_date(fixing_inx).fixing_date Thenn        n            n            If deal_ticket.get_fixing_date(fixing_inx).fixing_date > div_schedule.get_dividend(div_inx).dividend_date Thenn                time_step_ = div_schedule.get_dividend(div_inx).dividend_date - date_(inx)n                dt_(inx) = time_step_ / days_per_yearn            Elsen                time_step_ = deal_ticket.get_fixing_date(fixing_inx).fixing_date - date_(inx)n                dt_(inx) = time_step_ / days_per_yearn            End Ifn            n        Elsen            n            If date_(inx) + deal_ticket.time_step > div_schedule.get_dividend(div_inx).dividend_date Thenn                time_step_ = div_schedule.get_dividend(div_inx).dividend_date - date_(inx)n                dt_(inx) = time_step_ / days_per_yearn            Elsen                time_step_ = deal_ticket.time_stepn                dt_(inx) = time_step_ / days_per_yearn            End Ifn            n        End Ifn        n    End Ifn    n    inx = inx + 1n    n    ReDim Preserve date_(0 To inx) As Daten    ReDim Preserve dt_(0 To inx) As Doublen    ReDim Preserve is_fixing_date_(0 To inx) As Booleann    ReDim Preserve is_first_fixing_date_(0 To inx) As Booleann    ReDim Preserve is_dividend_date_(0 To inx) As Booleann    n    date_(inx) = date_(inx - 1) + time_step_n    n        n    Exit Subn    nErrorHandler:nn    Err.Raise vbObjectError + 1002, "clsTimeStep.make_next_step :" & Chr(13) & Err.source, Err.description '1002 <-- initialization errorn    Exit Subn    nnEnd Subnn'------------------------------------------------n' called by Initializen'------------------------------------------------nPrivate Sub fixing_check(inx As Integer, ByRef fixing_inx As Integer, ByRef first_fixing_checked As Boolean, deal_ticket As clsCliquetDealTicket)nn     nOn Error GoTo ErrorHandlern    n    '---------------------------------------------------------------------------n    ' If fixing index is larger than the boundary, which is the case when the current date is later than last fixing daten    n    If fixing_inx <= deal_ticket.get_max_fixing_index Thenn    n            n        n        If date_(inx) = deal_ticket.get_fixing_date(fixing_inx).fixing_date Thenn        n            is_fixing_date_(inx) = Truen            fixing_inx = fixing_inx + 1n            n            If first_fixing_checked Thenn            n                is_first_fixing_date_(inx) = Falsen                n            Elsen                n                is_first_fixing_date_(inx) = Truen                first_fixing_checked = Truen                n                n            End Ifn            n        Elsen        n            is_fixing_date_(inx) = Falsen            is_first_fixing_date_(inx) = Falsen            n        End Ifn        n    End Ifn        n    Exit Subn    nErrorHandler:nn    Err.Raise vbObjectError + 1002, "clsTimeStep.fixing_check :" & Chr(13) & Err.source, Err.description '1002 <-- initialization errorn    Exit Subn    nnEnd Subnn'------------------------------------------------n' called by Initializen'------------------------------------------------nPrivate Sub div_check(inx As Integer, ByRef div_inx As Integer, div_schedule As clsDividendSchedule)n    nOn Error GoTo ErrorHandlern n    If date_(inx) = div_schedule.get_dividend(div_inx).dividend_date Thenn    n        is_dividend_date_(inx) = Truen        div_inx = div_inx + 1n               n    Elsen    n        is_dividend_date_(inx) = Falsen        n    End Ifn        n    Exit Subn    nErrorHandler:nn    Err.Raise vbObjectError + 1002, "clsTimeStep.div_check :" & Chr(13) & Err.source, Err.description '1002 <-- initialization errorn    Exit Subn    nnEnd Sub