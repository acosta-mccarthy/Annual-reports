SELECT
LEFT(location_code, 3) AS "LOCATION",


--PART I - Total holdings by location and statistical code
COUNT(CASE when icode2 = 'a' then 1 end)AS "Adult Fiction 2.1",
COUNT(CASE when icode2 = 'b' then 1 end)AS "Adult Non-Fiction 2.2",
COUNT(CASE when icode2 = 'a' OR icode2 = 'b' then 1 end)AS "Total Adult 2.3",
COUNT(CASE when icode2 = 'c' then 1 end)AS "Juvenile Fiction 2.4",
COUNT(CASE when icode2 = 'd' then 1 end)AS "Juvenile Non-Fiction 2.5",
COUNT(CASE when icode2 = 'c' OR icode2 = 'd' then 1 end)AS "Total Juvenile Books 2.6",
COUNT(CASE when icode2 = 'a' OR icode2 = 'b' OR icode2 = 'c' OR icode2 = 'd' then 1 end)AS "Total Books 2.7",
COUNT(CASE when icode2 = 'f' then 1 end)AS "Microform",
COUNT(CASE when icode2 = 'h' then 1 end)AS "Adult Sound Recording",
COUNT(CASE when icode2 = 'i' then 1 end)AS "Adult Videorecording",
COUNT(CASE when icode2 = 'j' then 1 end)AS "Media",
COUNT(CASE when icode2 = 'l' then 1 end)AS "Adult Software",
COUNT(CASE when icode2 = 'm' then 1 end)AS "Equipment/Relia",
COUNT(CASE when icode2 = 'n' then 1 end)AS "Supressed Item",
COUNT(CASE when icode2 = 'q' then 1 end)AS "Juvenile Video",
COUNT(CASE when icode2 = 'r' then 1 end)AS "Juvenile Audio",
COUNT(CASE when icode2 = 's' then 1 end)AS "Juvenile Other Media",
COUNT(CASE when icode2 = 't' then 1 end)AS "Juvenile Software",
COUNT(CASE when icode2 = 'z' then 1 end)AS "Vertical File",
COUNT(CASE when icode2 = 'n' OR icode2 = 'z' then 1 end)AS "All Other Print 2.10",
COUNT(CASE when icode2 = 'a' OR icode2 = 'b' OR icode2 = 'c' OR icode2 = 'd' OR icode2 = 'n' OR icode2 = 'z' then 1 end)AS "Total Print 2.12",
--AS "eBook 2.13",
--AS "Audio Downloadable Units 2.17",
--AS "Total Videorecording Downloadable",
COUNT(CASE when icode2 = 'l' OR icode2 = 't' then 1 end)AS "Total Other Electronic Materials 2.19",
COUNT(CASE when icode2 = 'h' OR icode2 = 'r' then 1 end)AS "Total Sound Recording 2.21",
COUNT(CASE when icode2 = 'i' OR icode2 = 'q' then 1 end)AS "Total Videorecording 2.22",
COUNT(CASE when icode2 = 'f' OR icode2 = 'j' OR icode2 = 'm' OR icode2 = 's' then 1 end)AS "All Other Materials 2.23",
COUNT(CASE when icode2 = 'h' OR icode2 = 'r' OR icode2 = 'i' OR icode2 = 'q' OR icode2 = 'f' OR icode2 = 'j' OR icode2 = 'm' OR icode2 = 's' then 1 end)AS "Total Other Materials",
--AS "Grand Total Holdings",

--PART II - Total additions (all holdings added during previous year) by location and statistical code
COUNT(CASE when icode2 = 'a' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Fiction Added",
COUNT(CASE when icode2 = 'b' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Non-Fiction Added",
COUNT(CASE when icode2 = 'c' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Fiction Added",
COUNT(CASE when icode2 = 'd' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Non-Fiction Added",
COUNT(CASE when (icode2 = 'a' OR icode2 = 'b' OR icode2 = 'c' OR icode2 = 'd') AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Cataloged Books added 2.27",
COUNT(CASE when icode2 = 'l' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Software Added",
COUNT(CASE when icode2 = 't' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Software Added",
--AS "eBooks Added 2.29"
--AS "Electronic Materials Added"
COUNT(CASE when icode2 = 'f' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Microfilm Added",
COUNT(CASE when icode2 = 'h' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Sound Recording Added",
COUNT(CASE when icode2 = 'i' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Videorecording Added",
COUNT(CASE when icode2 = 'j' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Adult Other Media Added",
COUNT(CASE when icode2 = 'm' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Equipment/Realia Added",
COUNT(CASE when icode2 = 'n' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Suppressed Items Added",
COUNT(CASE when icode2 = 'q' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Videorecording Added",
COUNT(CASE when icode2 = 'r' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Sound Recording Added",
COUNT(CASE when icode2 = 's' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Juvenile Other Media Added",
COUNT(CASE when icode2 = 'z' AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "Vertical File Added",
-- AS "Other Media Added",
--AS "Downloadable Audio Added"
COUNT(CASE when (icode2 = 'n' OR icode2 = 'z') AND record_creation_date_gmt >=  DATE_TRUNC('day', now()) - interval '1 year' AND record_creation_date_gmt < DATE(NOW()) then 1 end) AS "All Other Print Materials Added 2.28"
-- AS "All other materials added"
-- AS "Total Added"


FROM
sierra_view.item_view

WHERE
location_code LIKE 'bea%' OR
location_code LIKE 'cld%' OR
location_code LIKE 'hil%' OR
location_code LIKE 'mah%' OR
location_code LIKE 'mar%'
--Limits locations to school district libraries only

GROUP BY "LOCATION"
ORDER BY "LOCATION"