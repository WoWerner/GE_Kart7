select count (kommunionen.PersonenID) as Anzahl,
       personen.Vorname,
       personen.Nachname
from kommunionen
left outer join personen on personen.PersonenID = kommunionen.PersonenID
where strftime('%Y',kommunionen.AbendmahlsDatum) = strftime('%Y','now')
Group by personen.PersonenID
order by personen.Nachname, personen.Vorname