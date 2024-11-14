% Data import & cleaning    
sp_daily_prices = readtable("C:\Users\Ian Burger\Numoro\Numoro Team - Documents\Ian\Excel\Morningstar sheets\RandomS&PPrices Daily 2.xlsx", ...
    "Sheet","Data Sheet", ...
    "NumHeaderLines", 4); 

% Drop the last column to get rid of most NaNs
sp_daily_prices(:, end) = [];

% Extract the first row as variable names
variableNames = sp_daily_prices(1, :);

% Drop variable names from table
sp_daily_prices(1,:) = [];

% Convert variableNamesCell to cell array
variableNamesCell = table2cell(variableNames);

% Convert the numerical date strings to datetime format (from Excel serial date)
for i = 3:length(variableNamesCell)
    % Convert each date number to datetime
    variableNamesCell{i} = datestr(datetime(variableNamesCell{i}, 'ConvertFrom', 'excel'), 'yyyymmdd');
end

% Assign the new variable names to the price data table
sp_daily_prices.Properties.VariableNames = variableNamesCell;

% Convert date columns (assumed to be in the 3rd column onward) to datetime
dateColumn = datetime(sp_daily_prices.Properties.VariableNames(3:end), 'InputFormat', 'yyyyMMdd');

% Extract the price data starting from the 3rd column onward
priceData_daily = sp_daily_prices{:, 3:end};
rowsWithNaN = any(isnan(priceData_daily), 2); 
priceData_daily = priceData_daily(~rowsWithNaN, :);
sp_daily_prices = sp_daily_prices(~rowsWithNaN, :);

daily_returns = price2ret(priceData_daily'); 
daily_returns = daily_returns';

returnDailyColumnNames = sp_daily_prices.Properties.VariableNames(4:end);
sp_daily_returnsTable = sp_daily_prices(:,1:2);
sp_daily_returnsTable = [sp_daily_returnsTable, array2table(daily_returns)];
sp_daily_returnsTable.Properties.VariableNames(3:end) = returnDailyColumnNames;

% Filter prices that only fall on Fridays
isFriday = (weekday(dateColumn) == 6);  % MATLAB weekday function: 6 corresponds to Friday

% Extract only Friday prices
sp_weekly_prices = sp_daily_prices(:, [1, 2, find(isFriday) + 2]);  % Include stock ID and stock names

% Filter out rows (stocks) that contain any NaN values in the weekly prices
priceData_weekly = sp_weekly_prices{:, 3:end}; 
rowsWithNaNw = any(isnan(priceData_weekly), 2); 
filtered_sp_weekly_prices = sp_weekly_prices(~rowsWithNaNw, :);

% Compute weekly returns from filtered weekly prices using price2ret
% Here, we assume prices start from column 3 onward
weekly_prices = filtered_sp_weekly_prices{:, 3:end};
weekly_returns = price2ret(weekly_prices');

% Transpose returns to match original stock row format
weekly_returns = weekly_returns';

% Create a new table for weekly returns
sp_weekly_returnsTable = filtered_sp_weekly_prices(:, 1:2);  % Keep first two columns (stock IDs and names)
sp_weekly_returnsTable = [sp_weekly_returnsTable, array2table(weekly_returns)];  % Append weekly returns

% Assign the correct column names for the weekly returns (from the date columns)
returnColumnNames = sp_weekly_prices.Properties.VariableNames(4:end);  % Get column names for the returns
sp_weekly_returnsTable.Properties.VariableNames(3:end) = returnColumnNames;  % Assign to return table