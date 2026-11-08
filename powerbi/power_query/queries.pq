// =====================================================================
// Power Query (M). One block = one query. Paste each into
// Home > Transform data > New Source > Blank Query > Advanced Editor.
//
// First create a Text parameter  ProjectFolder = C:\...\sales-ops-monthly-reporting\
// (with the trailing backslash). Run `python run_all.py` beforehand so
// data/processed/ exists.
// =====================================================================


// ---------- query: fact_sales_line  (~1.03M rows)
let
    Source  = Csv.Document(File.Contents(ProjectFolder & "data\processed\fact_sales_line.csv"),
                           [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed   = Table.TransformColumnTypes(Headers, {
        {"invoice", type text}, {"invoice_ts", type datetime}, {"invoice_date", type date},
        {"month", type date}, {"stock_code", type text}, {"quantity", Int64.Type},
        {"unit_price", type number}, {"line_revenue", Currency.Type}, {"customer_id", Int64.Type},
        {"is_guest", Int64.Type}, {"country", type text}, {"line_type", type text},
        {"source_sheet", type text}}, "en-US"),
    // timestamp and source sheet are not needed in the model - smaller, faster file
    Slim    = Table.RemoveColumns(Typed, {"invoice_ts", "source_sheet"})
in
    Slim


// ---------- query: dim_product
let
    Source  = Csv.Document(File.Contents(ProjectFolder & "data\processed\dim_product.csv"),
                           [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed   = Table.TransformColumnTypes(Headers, {
        {"stock_code", type text}, {"median_unit_price", type number},
        {"first_sold", type date}, {"description", type text}}, "en-US"),
    Proper  = Table.TransformColumns(Typed, {{"description", Text.Proper, type text}})
in
    Proper


// ---------- query: dim_customer
let
    Source  = Csv.Document(File.Contents(ProjectFolder & "data\processed\dim_customer.csv"),
                           [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed   = Table.TransformColumnTypes(Headers, {
        {"customer_id", Int64.Type}, {"first_purchase", type date}, {"last_purchase", type date},
        {"home_country", type text}, {"cohort_month", type date}, {"region", type text}}, "en-US")
in
    Typed


// ---------- query: dim_country
let
    Source  = Csv.Document(File.Contents(ProjectFolder & "data\processed\dim_country.csv"),
                           [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    Headers


// ---------- query: data_quality_log  (shown on the Method page)
let
    Source  = Csv.Document(File.Contents(ProjectFolder & "reports\data_quality_log.csv"),
                           [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    Typed   = Table.TransformColumnTypes(Headers, {{"rows", Int64.Type}}, "en-US")
in
    Typed
