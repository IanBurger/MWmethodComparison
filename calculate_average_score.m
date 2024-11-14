function tb_ave_score = calculate_average_score(portfolioWeeklyPrices, portInfo)    
    
    % Reverse weekly price order
    currentOrder = portfolioWeeklyPrices.Properties.VariableNames;
    revOrder = fliplr(currentOrder);
    portfolioWeeklyPrices = portfolioWeeklyPrices(:,revOrder);
    
    % Extract price array from weekly prices
    pricesArray = table2array(portfolioWeeklyPrices(:,:));
    
    [nStocks, nWeeks] = size(pricesArray);
    
    returnsMatrix = zeros(nStocks, nWeeks-1);
    
    for i=1:nWeeks-1
        returnsMatrix(:, i) = (pricesArray(:, i)./ pricesArray(:, i+1)-1);
    end
    
    % Calculate quality using the coefficient of variation (CV)
    % CV = mean return / standard deviation of returns (used as a risk metric)
    stdDevReturns = std(returnsMatrix, 0, 2);  % Standard deviation of returns for each stock
    meanReturns = mean(returnsMatrix, 2);      % Average return for each stock
    
    % Calculate the quality measure based on CV^(-1)
    qualityRaw = meanReturns ./ stdDevReturns;
    qualityNormalized = normalize_range(qualityRaw);  % Normalizing quality to 0-100 range
    
    % Step 4: Create table for returns and add stock information
    tb_returns = array2table(returnsMatrix, 'VariableNames', portfolioWeeklyPrices.Properties.VariableNames(1:46));
    tb_returns.SecID = portInfo.SecID;
    tb_returns.Name = portInfo.Name;
    tb_returns.Quality = qualityNormalized;
    tb_returns.MeanReturns = meanReturns;
    tb_returns.StdDevReturns = stdDevReturns;
    
    % Step 5: Calculate strength and consistency using 13-week and 26-week lookbacks
    lookbackStrength13Weeks = calculate_lookback_strength(pricesArray, 13, 0.7); % 13-week lookback with 70% recency bias
    lookbackStrength26Weeks = calculate_lookback_strength(pricesArray, 26, 0.3); % 26-week lookback with 30% recency bias
    
    % Combine both lookbacks into a blended strength score
    blendedStrength = lookbackStrength13Weeks + lookbackStrength26Weeks;
    
    % Step 6: Calculate the strength score for each stock (scaling to 0-100 range)
    strengthMatrix = scale_to_100(blendedStrength);
    
    % Step 7: Compute weighted average strength using predefined weights
    weights = [2, 2, 1.75, 1.75, 1.5, 1.5, 1.25, 1.25, 1, 1];
    weightedStrength = compute_weighted_average(strengthMatrix, weights);
    
    % Normalize the weighted strength score to 0-100 range
    consistencyScore = normalize_range(weightedStrength);
    
    % Step 8: Combine the Quality, Strength, and Consistency into a summary table
    averageScore = mean([weightedStrength, consistencyScore, qualityNormalized], 2);
    tb_summary = table(weightedStrength, consistencyScore, qualityNormalized, averageScore, ...
        'VariableNames', {'Strength', 'Consistency', 'Quality', 'AverageScore'});
    tb_summary.SecID = tb_returns.SecID;
    tb_summary.Name = tb_returns.Name;
    
    % Step 9: Normalize average score and apply exponential weighting
    normAverageScore = tb_summary.AverageScore / max(tb_summary.AverageScore);
    % expCoeff = 0.2;
    % expWeight = normAverageScore .^ expCoeff;
    % finalWeights = 100 * (expWeight / sum(expWeight));  % Normalize final weights to sum up to 100%
    
    % Step 10: Create final output table
    tb_ave_score = table(tb_summary.SecID, tb_summary.Name, tb_summary.AverageScore, ...
        tb_summary.Strength, tb_summary.Consistency, tb_summary.Quality, normAverageScore,  ...
        'VariableNames', {'SecID', 'Name', 'AverageScore', 'Strength', 'Consistency', 'Quality', ...
        'NormalizedScore'});   
end

%% Helper Functions

% Normalization to scale values to the range 0-100
function normalizedValues = normalize_range(values)
    minVal = min(values);
    maxVal = max(values);
    normalizedValues = 100 * (values - minVal) / (maxVal - minVal);
end

% Function to calculate lookback strength for a given period (lookbackWeeks)
function lookbackStrength = calculate_lookback_strength(priceData, lookbackWeeks, biasFactor)
    [numStocks, ~] = size(priceData);
    lookbackStrength = zeros(numStocks, 10);  % Initialize lookback strength matrix
    
    for i = 1:10
        startWeek = i;
        lookbackWeek = i + lookbackWeeks - 1;
        lookbackStrength(:, i) = biasFactor * ((priceData(:, startWeek) - priceData(:, lookbackWeek)) ./ priceData(:, lookbackWeek));
    end
end

% Function to scale the strength matrix to the range of 0-100
function scaledMatrix = scale_to_100(matrix)
    minVal = min(matrix);
    maxVal = max(matrix);
    scaledMatrix = 100 * (matrix - minVal) ./ (maxVal - minVal);
end

% Function to compute the weighted average of strength scores
function weightedAverage = compute_weighted_average(strengthMatrix, weights)
    weightedScores = strengthMatrix .* repmat(weights, size(strengthMatrix, 1), 1);  % Apply weights
    weightedAverage = sum(weightedScores, 2) / sum(weights);  % Calculate weighted average
end
