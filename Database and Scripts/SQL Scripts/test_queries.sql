-- Output - IDR Fund Level
USE idr_db

-- Diversification
-- •	Total – All Regions: should just be the total NAV of the appropriate fund/account from the IDR Quarterly Inputs tab for Value Type 2
select netassetvalue
from idrquarterlyinputs
where cfid = 1
and quarterid = 1

-- •	All Other Rows (US Northeast through Other)
-- o	Value Type 1: Input - CF Diversification (NAV) tab lists the breakouts by percentage for region/property type, where each component fund will have a separate 
-- tab like this (cells A1:O11)
(select totalamt
from diversification
where cfid = 1
and quarterid = 1
and fieldname = 'US Northeast')/
(select netassetvalue
from idrquarterlyinputs
where cfid = 1
and quarterid = 1) -- and so on

select office
from diversification
where cfid = 1
and quarterid = 1
and fieldname = 'Total (all regions)'/
(select netassetvalue
from idrquarterlyinputs
where cfid = 1
and quarterid = 1) -- and so on


-- 	Weight the funds by weighting for appropriate NAV for each CF column in IDR Quarterly Inputs tab, by the respective region/property type exposure
-- 	Same for $ totals – this can just be the weighted percentage in Value Type 1 column multiplied by total NAV in cell C3

-- Life Cycle
-- •	Same as diversification, except using cells A12:O21 in the Input – CF Diversification tab
-- Structure
-- •	Same as diversification, except using cells A22:O25 in the Input – CF Diversification tab
-- Valuations & Other
-- •	Same as diversification, except using cells A26:O35 in the Input – CF Diversification tab
-- Fund Level Data
-- •	Number of Investments to Net Investor Cash Flow, and Investment Queue to Forward Commitments (cells B49-B64, B68-B73) – just add the data together
-- from the Input – CF Fund Level (NAV) tab for any component fund that has an allocation in the overall fund from the IDR Quarterly Inputs tab (allocation >$0)

select value1
from fundlevelnav
where fieldname = 'Number of Investments'
and cfid = 1
and quarterid = 1 -- and so on

-- •	Dividend Yield, Cash Income Yield, NOI Yield (cells B65-B67) - weight the funds by weighting for appropriate NAV for each CF column in 
-- IDR Quarterly Inputs tab, by the respective field in the Input – CF Fund Level (NAV) tab
(select value
from fundlevelnav
where fieldname = 'Number of Investments'
and cfid = 1
and quarterid = 1) /
(select netassetvalue
from idrquarterlyinputs
where cfid = 1
and quarterid = 1) -- and so on


-- Leverage
-- •	Same methodology for all fields in this column as Dividend Yield, NOI Yield in Fund Level Data section above


-- Time Weighted Returns
-- •	Weight the funds by weighting for appropriate NAV for each CF column in IDR Quarterly Inputs tab, by the respective field in the Input – CF Performance tab

-- Acquisitions, Dispositions, Portfolio, Disclosures Sections
-- •	We don’t need these for the aggregate – we will need to be able to query this in an organized manner in the new interface
