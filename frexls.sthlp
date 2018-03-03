{smcl}
{* *! version 1.0  nov2017}{...}
{cmd:help frexls}
{hline}

{title:Title}

{phang}
{cmd:frexls} {hline 2} esporta in excel l'output del comando {stata "ssc desc fre":fre}. {cmd fre} non è un comando ufficiale di Stata, ma per usare {cmd frexls} deve essere installato.


{title:Syntax}

{p 8 14 2}
{cmd:frexls} {it:varname} {if} [{cmd:,} {it:fre options}] [{it:export excel options}]

{pstd}
dove bal bla bla



{title:Description}

{pstd}
A number of {cmd:egen} functions 
(both built-in and user-written, e.g. {stata "ssc desc egenmore":egenmore})
expect a {help varlist:{it:varlist}}
and perform observation-level calculations based on
the variables provided as argument.
With {cmd:tsegen}, you can invoke any of these {bf:egen} functions
using a time-series varlist ({help tsvarlist:{it:tsvarlist}})
instead.
{cmd:tsegen} converts the {help tsvarlist:{it:tsvarlist}} 
to a {help varlist:{it:varlist}} 
by substituting equivalent temporary variables as necessary
and then invokes the specified {help egen:{bf:egen}} function.



{title:Options}
{dlgtab:fre options}

{synopt :{opt nonmis:sing}} specifica che i valori missing di {it:varname} non devono essere considerati.{p_end}

{synopt :{opt as:cending}} specifica che le modalità di {it:varname} vengano visualizzate in ordine crescente.{p_end}

{synopt :{opt de:scending}} specifica che le modalità di {it:varname} vengano visualizzate in ordine decrescente.{p_end}

{synopt :{opt all}} specifica che tutte le modalità di {it:varname} vengano visualizzate.{p_end}

{synopt :{opt nov:alue}} specifica che il valore numerico di {it:varname} non venga visualizzato, ma solo la descrizione.{p_end}




{marker output}{...}
{dlgtab:export excel options}

{synopt :{opt xlsfile(filename.ext)}} specifica il nome del file excel dove salvare l'output. Deve essere specificata l'estensione e può essere indicato anche un percorso.{p_end}

{synopt :{opt sheet(str)}} specifica il nome del foglio dove salvare l'output.{p_end}

{synopt:{opt replace|sheetreplace|sheetmodify}} bisogna scegliere una di queste tre opzioni. {opt replace} crea un nuovo file excel, {opt sheetreplace} crea un nuovo foglio o sostituisce uno già esistente ma non modifica il 
resto del file, {opt sheetmodify} modifica un foglio già esistente.{p_end}

{synopt :{opt cell($#)}} specifica la cella da cui iniziare per scrivere la tabella. Se non specificata di default si assume la cella A1.{p_end}

{synopt :{opt title(string)}} specifica l'eventuale intestazione da mettere prima della tabella.{p_end}

{synopt :{opt post(string)}} specifica l'eventuale testo da mettere alla fine della tabella.{p_end}

{synopt :{opt intr_size(number)}} specifica la larghezza della prima colonna, quella che contiene la descrizione delle modalità di {it:varname}. Il valore di default è 45.{p_end}

{synopt :{opt res_size(number)}} specifica la larghezza delle colonne successive alla prima. Il valore di default è 11.{p_end}

{synopt :{opt fontname(string)}} specifica il nome del font da usare. Il valore di default è Calibri.{p_end}

{synopt :{opt fontsize(number)}} specifica la dimensione del font. Il valore di default è 11.{p_end}

{synopt :{opt pattern_intc(string)}} specifica il colore di sfondo per l'intestazione delle colonne. 
E' possibile usare il formato RGB specificandolo così {opt pattern_intc("6 32 83")} oppure una lista di colori. 
Vedi in fondo per la lista dei colori. Il default è nessun colore.{p_end}











{title:Colors}

	{cmd:aliceblue}
	{cmd:antiquewhite}
	{cmd:aqua}
	{cmd:aquamarine}
	{cmd:azure}
	{cmd:beige}
	{cmd:bisque}
	{cmd:black}
	{cmd:blanchedalmond}
	{cmd:blue}
	{cmd:blueviolet}
	{cmd:brown}
	{cmd:burlywood}
	{cmd:cadetblue}
	{cmd:chartreuse}
	{cmd:chocolate}
	{cmd:coral}
	{cmd:cornflowerblue}
	{cmd:cornsilk}
	{cmd:crimson}
	{cmd:cyan}
	{cmd:darkblue}
	{cmd:darkcyan}
	{cmd:darkgoldenrod}
	{cmd:darkgray}
	{cmd:darkgreen}
	{cmd:darkkhaki}
	{cmd:darkmagenta}
	{cmd:darkolivegreen}
	{cmd:darkorange}
	{cmd:darkorchid}
	{cmd:darkred}
	{cmd:darksalmon}
	{cmd:darkseagreen}
	{cmd:darkslateblue}
	{cmd:darkslategray}
	{cmd:darkturquoise}
	{cmd:darkviolet}
	{cmd:deeppink}
	{cmd:deepskyblue}
	{cmd:dimgray}
	{cmd:dodgerblue}
	{cmd:firebrick}
	{cmd:floralwhite}
	{cmd:forestgreen}
	{cmd:fuchsia}
	{cmd:gainsboro}
	{cmd:ghostwhite}
	{cmd:gold}
	{cmd:goldenrod}
	{cmd:gray}
	{cmd:green}
	{cmd:greenyellow}
	{cmd:honeydew}
	{cmd:hotpink}
	{cmd:indianred }
	{cmd:indigo }
	{cmd:ivory}
	{cmd:khaki}
	{cmd:lavender}
	{cmd:lavenderblush}
	{cmd:lawngreen}
	{cmd:lemonchiffon}
	{cmd:lightblue}
	{cmd:lightcoral}
	{cmd:lightcyan}
	{cmd:lightgoldenrodyellow}
	{cmd:lightgray}
	{cmd:lightgreen}
	{cmd:lightpink}
	{cmd:lightsalmon}
	{cmd:lightseagreen}
	{cmd:lightskyblue}
	{cmd:lightslategray}
	{cmd:lightsteelblue}
	{cmd:lightyellow}
	{cmd:lime}
	{cmd:limegreen}
	{cmd:linen}
	{cmd:magenta}
	{cmd:maroon}
	{cmd:mediumaquamarine}
	{cmd:mediumblue}
	{cmd:mediumorchid}
	{cmd:mediumpurple}
	{cmd:mediumseagreen}
	{cmd:mediumslateblue}
	{cmd:mediumspringgreen}
	{cmd:mediumturquoise}
	{cmd:mediumvioletred}
	{cmd:midnightblue}
	{cmd:mintcream}
	{cmd:mistyrose}
	{cmd:moccasin}
	{cmd:navajowhite}
	{cmd:navy}
	{cmd:oldlace}
	{cmd:olive}
	{cmd:olivedrab}
	{cmd:orange}
	{cmd:orangered}
	{cmd:orchid}
	{cmd:palegoldenrod}
	{cmd:palegreen}
	{cmd:paleturquoise}
	{cmd:palevioletred}
	{cmd:papayawhip}
	{cmd:peachpuff}
	{cmd:peru}
	{cmd:pink}
	{cmd:plum}
	{cmd:powderblue}
	{cmd:purple}
	{cmd:red}
	{cmd:rosybrown}
	{cmd:royalblue}
	{cmd:saddlebrown}
	{cmd:salmon}
	{cmd:sandybrown}
	{cmd:seagreen}
	{cmd:seashell}
	{cmd:sienna}
	{cmd:silver}
	{cmd:skyblue}
	{cmd:slateblue}
	{cmd:slategray}
	{cmd:snow}
	{cmd:springgreen}
	{cmd:steelblue}
	{cmd:tan}
	{cmd:teal}
	{cmd:thistle}
	{cmd:tomato}
	{cmd:turquoise}
	{cmd:violet}
	{cmd:wheat}
	{cmd:white}
	{cmd:whitesmoke}
	{cmd:yellow}
	{cmd:yellowgreen}



    
{title:Examples}

{pstd}
Calculate the mean of the variable {hi:invest} over a 5-year 
rolling window that includes the current observation

        {cmd:.} {stata webuse grunfeld, clear}
        {cmd:.} {stata tsegen inv_m5 = rowmean(invest L(1/4).invest)}


{pstd}
Thanks to Sebastian Kripfganz for pointing this out on
{browse "http://www.statalist.org/forums/forum/general-stata-discussion/general/1292241-new-on-ssc-tsegen-for-computations-over-a-rolling-window-using-time-series-operators-with-egen-functions?p=1292371#post1292371":Statalist}.


{title:Limitations}

{pstd}


{title:Author}

{pstd}Nicola Tommasi{p_end}
{pstd}nicola.tommasi@univr.it{p_end}


{marker also}{...}
{title:Also see}

{psee}
Stata:  {help xl()} - {help fre} if installed
{p_end}

{psee}
Article: Chuck Huber (2017) {browse "https://blog.stata.com/2017/04/06/":Creating Excel tables with putexcel part 3: Writing custom reports for arbitrary variables}, {it:The Stata Blog}

Article: Chuck Huber (2017) {browse "https://blog.stata.com/2017/01/24/":Creating Excel tables with putexcel, part 2: Macro, picture, matrix, and formula expressions}, {it:The Stata Blog}

Article: Chuck Huber (2017) {browse "https://blog.stata.com/2017/01/10/creating-excel-tables-with-putexcel-part-1-introduction-and-formatting/": Creating Excel tables with putexcel, part 1: Introduction and formatting}, {it:The Stata Blog} 

Article:  Kevin Crow (2013) {browse "https://blog.stata.com/2013/09/25/export-tables-to-excel/":Export tables to Excel}, {it:The Stata Blog} 

