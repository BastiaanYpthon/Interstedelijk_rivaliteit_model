Stappen naar een model maken waarmee de aannemelijkheid van rivaliteit tussen twee steden kan worden berekend.

Tijdens het schrijven van mijn Bachelor Scriptie voor Sociale Geografie en Planologie heb ik doormiddel van verschillende codes een dataset gemaakt. Deze dataset en codes om onder andere de dataset te maken, maar ook codes waarmee bepaalde analyses ter validatie kunnen worden uitgevoerd zullen hier te vinden zijn met waar nodig toegelichting. Toelichting voor de verschillende data en codes:

Voor 57 verschillende steden is op het internet van openbare bronnen, data verzameld die samen een dataset vormen. Deze dataset is steden_data.xlsx, vervolgens wordt deze dataset door de code steden_paren_met_cooccurence.py verwerkt naar stad_data_output.xlsx, wat een matrix is waar alle stedenparen worden vergeleken.
Hier zijn 1597 vergelijkingen. Om van deze data waardes te maken is de code data_normalized.py gebruikt waaruit normalized_rivalry.xlsx komt, waar alle waardes naar genormaliseerde scores tussen 0 en 1 zijn berekend.
Vervolgens zijn deze scores aan de hand van gewichten die uit een enquête zijn gekomen aangepast door weighted_rival.py te runnen, waar weighted_rivalry_output.xlsx uit komt.
In dat document zijn de scores te zien per factoor in het tweede tabblad en in de laatste kolom de totale gewogen score waarmee de rivaliteit aannemelijkheidsscore wordt uitgedrukt.
Door een selectie te maken van bekende rivaliteiten is het model en gecontroleerd, er wordt gekeken hoevaak deze bekende rivaliteiten in de top 10, top 25, top 50 en top 100 staat in de code bekende_in_toppen.py waar het document rivalen_positions.xlsx uit ontstaat.
De verschillende gewichten worden doormiddel van een weight variation analysis in weight_variations.py gerund waar weighted_rivalry_analysis_output.xlsx uit komt. 
Daarin zijn voor alle wegingen kleine variaties toegepast om te kijken wat dat met de posities van de geselecteerde rivaliteiten doet. 
Een andere test die toegepast kan worden is een excluding factors analysis, die in excluding_factors.py wordt uitgevoerd en leidt tot weighted_rivalry_exclusion_analysis_output.xlsx met wederom de posities van de geselecteerde rivaliteiten wanneer er factoren om en om niet worden gebruikt in het model. 
Tot slot zijn er in een cluster analyse clusters van factoren samengesteld om te kijken hoe de geselecteerde steden worden gescoort aan de hand van kleinere groepjes factoren. 
Dit gebeurt in cluster_analyse.py waaruit weighted_rivalry_cluster_analysis_output.xlsx ontstaat.
