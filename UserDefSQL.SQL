select count (kommunionen.PersonenID) as Abendmahlsgänge,
       personen.Vorname,
       personen.Nachname
from kommunionen
left outer join personen on personen.PersonenID = kommunionen.PersonenID
where strftime('%Y',kommunionen.AbendmahlsDatum) = strftime('%Y','now')
group by personen.PersonenID
order by personen.Nachname, personen.Vorname