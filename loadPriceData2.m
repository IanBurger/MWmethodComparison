% Data import & cleaning    
spDailyPricesTb = readtable("C:\Users\Ian Burger\Numoro\Numoro Team - Documents\Ian\Excel\Morningstar sheets\RandomS&PPrices Daily 2.xlsx", ...
    "Sheet","Data Sheet", ...
    "NumHeaderLines", 4); 

% Drop the last column to get rid of most NaNs
spDailyPricesTb(:, end) = [];

% Extract the first row as variable names
variableNames = spDailyPricesTb(1, :);

% Drop variable names from table
spDailyPricesTb(1,:) = [];

% Convert variableNamesCell to cell array
variableNamesCell = table2cell(variableNames);

% Convert the numerical date strings to datetime format (from Excel serial date)
for i = 3:length(variableNamesCell)
    % Convert each date number to datetime
    variableNamesCell{i} = datestr(datetime(variableNamesCell{i}, 'ConvertFrom', 'excel'), 'yyyymmdd');
end

% Assign the new variable names to the price data table
spDailyPricesTb.Properties.VariableNames = variableNamesCell;

% Convert date columns (assumed to be in the 3rd column onward) to datetime
dateColumns = datetime(spDailyPricesTb.Properties.VariableNames(3:end), 'InputFormat', 'yyyyMMdd');

% Extract the price data starting from the 3rd column onward
spDailyPricesAr = spDailyPricesTb{:, 3:end};
rowsWithNaN = any(isnan(spDailyPricesAr), 2); 
spDailyPricesTb = spDailyPricesTb(~rowsWithNaN, :);
spDailyPricesAr = spDailyPricesAr(~rowsWithNaN, :);

% Filter prices that only fall on Fridays
isFriday = (weekday(dateColumns) == 6);  % MATLAB weekday function: 6 corresponds to Friday

% Extract only Friday prices
spWeeklyPricesTb = spDailyPricesTb(:, [1, 2, find(isFriday) + 2]);  % Include stock ID and stock names

% Filter out rows (stocks) that contain any NaN values in the weekly prices
priceData_weekly = spWeeklyPricesTb{:, 3:end}; 
rowsWithNaNw = any(isnan(priceData_weekly), 2); 
spWeeklyPricesTb = spWeeklyPricesTb(~rowsWithNaNw, :);

% Here, we assume prices start from column 3 onward
weekly_pricesAr = spWeeklyPricesTb{:, 3:end};

% % Data import & cleaning    
% spDailyPricesTb = readtable("C:\Users\Ian Burger\Numoro\Numoro Team - Documents\Ian\Excel\Morningstar sheets\RandomS&PPrices Daily 2.xlsx", ...
%     "Sheet","Data Sheet", ...
%     "NumHeaderLines", 4); 
% 
% % Drop the last column to get rid of most NaNs
% spDailyPricesTb(:, end) = [];
% 
% % Extract the first row as variable names
% variableNames = spDailyPricesTb(1, :);
% 
% % Drop variable names from table
% spDailyPricesTb(1,:) = [];
% 
% % Convert variableNamesCell to cell array
% variableNamesCell = table2cell(variableNames);
% 
% % Convert the numerical date strings to datetime format (from Excel serial date)
% for i = 3:length(variableNamesCell)
%     % Convert each date number to datetime
%     variableNamesCell{i} = datestr(datetime(variableNamesCell{i}, 'ConvertFrom', 'excel'), 'yyyymmdd');
% end
% 
% % Assign the new variable names to the price data table
% spDailyPricesTb.Properties.VariableNames = variableNamesCell;
% 
% % Convert date columns (assumed to be in the 3rd column onward) to datetime
% dateColumns = datetime(spDailyPricesTb.Properties.VariableNames(3:end), 'InputFormat', 'yyyyMMdd');
% 
% % Extract the price data starting from the 3rd column onward
% spDailyPricesAr = spDailyPricesTb{:, 3:end};
% rowsWithNaN = any(isnan(spDailyPricesAr), 2); 
% spDailyPricesTb = spDailyPricesTb(~rowsWithNaN, :);
% spDailyPricesAr = spDailyPricesAr(~rowsWithNaN, :);
% 
% % Filter prices that only fall on Fridays
% isFriday = (weekday(dateColumns) == 6);  % MATLAB weekday function: 6 corresponds to Friday
% 
% % Extract only Friday prices
% spWeeklyPrices = spDailyPricesTb(:, [1, 2, find(isFriday) + 2]);  % Include stock ID and stock names
% 
% % Filter out rows (stocks) that contain any NaN values in the weekly prices
% priceData_weekly = spWeeklyPrices{:, 3:end}; 
% rowsWithNaNw = any(isnan(priceData_weekly), 2); 
% spWeeklyPrices = spWeeklyPrices(~rowsWithNaNw, :);
% 
% % Here, we assume prices start from column 3 onward
% weekly_prices = spWeeklyPrices{:, 3:end};