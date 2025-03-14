% This function reads Excel files with an unknown number of header rows,
% and regardless if the file has a unit row or not.

function [Table, Headers, Units, dataRowNo, headerRowNo, unitRowOrNot] ...
    = readExcel(File) 

% Determine the location of the header row: 
opts = detectImportOptions(File);
headerRowNo = str2double(extract(opts.VariableNamesRange, digitsPattern)); 

% Read all contents from the first worksheet of the Excel:  
ALL = readcell(File, 'Sheet', 1, 'DataRange', 'A1'); 

% Get the contents of the header row:
Headers = ALL(headerRowNo,:);

% In case there are missing fields within the headers, replace them with an
% unique text string. Note these strings must be unique. 
Ind1 = cellfun(@(X)any(ismissing(X),'all'), Headers);
Headers(Ind1) = cellstr("Missing_header_"+find(Ind1));

%% ------- Check to see if there is a unit row or not -------
% The assumption is that a unit row  should not contain any numerical values.

% The total number of numerical values in the potential unit row: 
Count_float = sum(cellfun(@isfloat, ALL(headerRowNo+1,:)));

if Count_float == 0  % Yes there is a unit row. 

    unitRowOrNot = 1;
    Units  = ALL(headerRowNo+1,:);
    dataValue = ALL(headerRowNo+2:end,:);
    dataRowNo = [1:size(dataValue,1)]' + headerRowNo+1;

    % Replace any <ismissing> units with empty text strings
    Ind2 = cellfun(@(X)any(ismissing(X),'all'), Units);
    Units(Ind2) = cellstr({' '});

else   % No unit row

    unitRowOrNot = 0;
    Units = repmat({' '},size(Headers));
    dataValue = ALL(headerRowNo+1:end,:);
    dataRowNo = [1:size(dataValue,1)]' + headerRowNo; 

end

%% --------- Remove any missing rows in the data ----------

% Identify rows with all columns missing: 
missingInd = ismissing(string(dataValue));
missingRow = sum(missingInd,2) == size(missingInd,2);

% Remove data rows with all missing values: 
dataValue = dataValue(~missingRow,:);  

% Remove the corresponding row numbers for these empty rows: 
dataRowNo = dataRowNo(~missingRow); 

%% Data columns could contain a mix of any of the below:
% (1) Floating numbers
% (2) Missing values
% (3) Empty strings
% (4) Non-empty strings that can be converted into doubles
% (5) Non-empty strings that can not be converted into doubles

%% Step One: Identify columns with at least one non-empty text string that can not 
% be converted into doubles and treat these columns as strings. 

% Cells with non-empty strings, including those that can potentially be 
% converted to doubles: 
map1 = cellfun(@(x)ischar(x) && ~isempty(deblank(x)),dataValue);

% Cells with strings that could potentially be converted to a double: 
map2 = ~isnan(str2double(dataValue));

% Columns with at least one non-empty string that can not be converted into doubles: 
tf = any(map1 & ~map2, 1);

% Read these columns as strings: 
dataValue(:, tf) = cellstr(string(dataValue(:, tf)));

%% Step Two: For the rest of the columns, convert any strings that can potentially 
% be convtered to a double into double
dataValue(~tf & map2) = num2cell(str2double(dataValue(~tf & map2)));

%% Step Three: For the rest of the columns, replace any empty strings and missing 
% values with -999. 

% Identify cells with empty strings, like ' ': 
map3 = cellfun(@(x)ischar(x) && isempty(deblank(x)),dataValue);

% Identify cells with missing values:
map4 = ismissing(string(dataValue));

% Replace these missing or  empty string values with -999:
dataValue(~tf & (map3 | map4)) = {-999};

Table = cell2table(dataValue);

%% Assign the Variable names and units to the table: 
Table.Properties.VariableNames = Headers;
Table.Properties.VariableUnits = Units;

end