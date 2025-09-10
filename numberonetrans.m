% Get the current script directory
scriptDir = fileparts(mfilename('fullpath'));

% Define relative paths
inputFilePath = fullfile(scriptDir, 'climate_change_dataset.xlsx');
outputDir = fullfile(scriptDir, 'MatlabExcel Assignment');

% Create output directory if it doesn't exist
if ~exist(outputDir, 'dir')
    mkdir(outputDir);
end

% Read the Excel file
T = readtable(inputFilePath, Range="A1:T54", ReadVariableNames=true);

%A = 2020
A = T(1:12,:);
%B = 2021
B = T(13:24,:);
%C = 2022
C = T(25:36,:);
%D = 2023
D = T(37:48,:);
%E = 2024
E = T(49:53,:);

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