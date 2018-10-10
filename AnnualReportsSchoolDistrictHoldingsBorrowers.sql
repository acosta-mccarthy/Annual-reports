SELECT

--PART III - Total borrowers by location and residency
LEFT(home_library_code, 3) AS "LOCATION",
COUNT(CASE when home_library_code IS NOT NULL AND ptype_code != '3' then 1 end)AS "Resident Borrowers 3.2",
COUNT(CASE when home_library_code IS NOT NULL AND ptype_code = '3' then 1 end)AS "Non-Resident Borrowers 3.3",
COUNT(CASE when home_library_code IS NOT NULL then 1 end)AS  "Total Number of Borrowers"

FROM
sierra_view.patron_view

WHERE
home_library_code LIKE 'bea%' OR
home_library_code LIKE 'cld%' OR
home_library_code LIKE 'hil%' OR
home_library_code LIKE 'mah%' OR
home_library_code LIKE 'mar%'
--Limits locations to school district libraries only

GROUP BY "LOCATION"
ORDER BY "LOCATION"