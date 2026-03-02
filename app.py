rcs = ""
for rid, d, ob, sv, qt, hr, un in recent:

    df = (
        f"<form method='POST' action='/produtividade/excluir/{rid}' style='display:inline;'>"
        f"<input type='hidden' name='_csrf' value='{csrf_token()}'>"
        f"<button class='btn-sm btn-del' onclick=\"return confirm('Excluir?')\">Excluir</button>"
        f"</form>"
    )

    ut = (
        f"<span class='mini' style='margin-left:auto;'>{un or ''}</span>"
        if is_admin()
        else ""
    )

    qts = fmt_num(qt, 2) or "—"
    hrs = fmt_num(hr, 1) or "—"

    rcs += (
        f"<div class='rc'>"
        f"<div class='rc-top'>"
        f"<div class='rc-obra'>{ob or ''}</div>"
        f"<div class='rc-date'>{d or ''}</div>"
        f"</div>"
        f"<div class='rc-srv'>{sv or ''}</div>"
        f"<div class='rc-nums'>"
        f"<div>Qtd: <strong>{qts}</strong></div>"
        f"<div>Horas: <strong>{hrs}</strong></div>"
        f"</div>"
        f"<div class='rc-act'>"
        f"<a class='btn-sm btn-edit' href='/produtividade/editar/{rid}'>Editar</a>"
        f"{df}"
        f"{ut}"
        f"</div>"
        f"</div>"
    )




