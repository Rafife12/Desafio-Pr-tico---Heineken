// ============================================================
// Consulta Power Query (linguagem M)
// Tratamento de dados — Desafio Heineken (Manutenção de Frota)
// ============================================================

let
    Origem = Excel.Workbook(File.Contents("Base de dados.xlsx"), null, true),
    #"Base de dados_Sheet" = Origem{[Item="Base de dados",Kind="Sheet"]}[Data],
    #"Cabeçalhos Promovidos" = Table.PromoteHeaders(#"Base de dados_Sheet", [PromoteAllScalars=true]),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{
        {"Data - Dia", Int64.Type},
        {"Nome do dia", type text},
        {"Mês", Int64.Type},
        {"Localidade", type text},
        {"MARCA", type text},
        {"TIPO CARROCERIA", type text},
        {"Placa", type text},
        {"Idade", Int64.Type},
        {"Valor (R$)", type number},
        {"Peça", type text}
    }),

    // Padronização da coluna TIPO CARROCERIA:
    // corrige a falta do prefixo "3/4" em "TRUCK BAÚ" (placa FUE4266),
    // sem afetar as linhas que já estavam corretas como "3/4 TRUCK BAÚ".
    // A categoria "TOCO BAÚ" (placa GAC4A48) NÃO é alterada — é um porte
    // de veículo legítimo e distinto, não um erro de digitação.
    #"Coluna Condicional Adicionada" = Table.AddColumn(#"Tipo Alterado", "TIPO CARROCERIA (Padronizado)", each
        if [TIPO CARROCERIA] = "TRUCK BAÚ" then "3/4 TRUCK BAÚ" else [TIPO CARROCERIA]
    ),
    #"Colunas Removidas" = Table.RemoveColumns(#"Coluna Condicional Adicionada",{"TIPO CARROCERIA"}),
    #"Colunas Renomeadas" = Table.RenameColumns(#"Colunas Removidas",{
        {"TIPO CARROCERIA (Padronizado)", "TIPO CARROCERIA"}
    }),

    // Sinalização das ordens de serviço com Valor (R$) = 0,
    // para confirmação posterior com o financeiro/oficina
    // (pode ser garantia — valor correto — ou falha de lançamento — subestima o gasto real)
    #"Coluna Condicional Adicionada1" = Table.AddColumn(#"Colunas Renomeadas", "Observação", each
        if [Valor (R$)] = 0 then "Confirmar com o financeiro ou com a oficina se é garantia ou falha de lançamento" else null
    ),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Coluna Condicional Adicionada1",{
        {"TIPO CARROCERIA", type text},
        {"Observação", type text}
    })
in
    #"Tipo Alterado1"
