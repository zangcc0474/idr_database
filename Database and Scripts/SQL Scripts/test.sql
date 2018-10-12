USE idr_db;

DELETE FROM quarter
WHERE quarter.name='Q42018';

Select * from acquisitions;

Select * from quarter;

select * from componentfund;

Select * from fundlevelnav;

fundlevelnav(SELECT MAX(QuarterID) + 1 FROM Quarter);

DELETE FROM Quarter WHERE name='Q32018';

DELETE FROM componentfund where name = 'CF 1';

Select * FROM IDRIpNFIODCE;

Select * FROM Disclosures;

Select * FROM IDRIpNFIODCEX;

Select * FROM Dispositions;