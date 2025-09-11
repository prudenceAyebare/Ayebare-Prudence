% Get the current script directory
scriptDir = fileparts(mfilename('fullpath'));

% Define relative paths
inputFilePath = fullfile(scriptDir, 'US_Stock_Data.xlsx');
outputDir = fullfile(scriptDir, 'MatlabExcel Assignment');

% Create output directory if it doesn't exist
if ~exist(outputDir, 'dir')
    mkdir(outputDir);
end

% Read the Excel file
T = readtable(inputFilePath, Range="A1:AM1014", ReadVariableNames=true);

% Extract data for each year
A = T(1:23,:);      % 2024
B = T(24:273,:);    % 2023
C = T(274:518,:);   % 2022
D = T(519:764,:);   % 2021
E = T(765:1013,:);  % 2020

% Convert to structure arrays
AS = table2struct(A);
BS = table2struct(B);
CS = table2struct(C);
DS = table2struct(D);
ES = table2struct(E);

% Convert back to tables and save
AT = struct2table(AS);
writetable(AT, fullfile(outputDir, 'AT.xlsx'));

BT = struct2table(BS);
writetable(BT, fullfile(outputDir, 'BT.xlsx'));

CT = struct2table(CS);
writetable(CT, fullfile(outputDir, 'CT.xlsx'));

DT = struct2table(DS);
writetable(DT, fullfile(outputDir, 'DT.xlsx'));

ET = struct2table(ES);
writetable(ET, fullfile(outputDir, 'ET.xlsx'));

disp('Processing complete. Files saved in MatlabExcel Assignment folder.');