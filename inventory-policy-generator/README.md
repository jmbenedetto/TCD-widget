# Inventory policy generator widget

Purpose:
- Run inside a Grist custom widget page bound to `Entrada_Politicas`.
- Let a user select one policy row and click one button to create persisted detailed rows in `SaidaDados_Politica`.

Technical stance:
- Trigger and execution both stay inside Grist.
- The widget requests `requiredAccess: "full"`.
- It uses the Grist plugin API (`grist.docApi` / `grist.raw.docApi`) to:
	- read `Entrada_Produtos`.
	- check existing rows in `SaidaDados_Politica`.
	- write `AddRecord` user actions directly back to the document.
	- optionally update generation-status fields on `Entrada_Politicas`.

Important limitation:
- This widget is written in HTML/CSS/JavaScript because Grist custom widgets run in the browser.
- A pure Python widget is not the standard Grist path for V01.
- If you want Python, it should be used as an offline/dev helper or external API worker, which is a different architecture.

Current row-generation assumption:
- V01 live build only needs these input fields per generated row:
	- `id_produto`.
	- `policy_ref`.
	- `bucket_planejamento`.
- Formula and lookup columns in `SaidaDados_Politica` fill the rest.

How to bind in Grist:
- Add a custom widget page.
- URL should point to the hosted `inventory-policy-generator/index.html`.
- Bind the widget to `Entrada_Politicas`.
- Grant `full` access.
- Select one policy row, then click `Generate detail rows`.
