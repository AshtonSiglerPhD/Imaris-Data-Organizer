%% Simple single parameter organizer for Imaris files
% Ashton L Sigler, January 23rd, 2023
% This script is designed to take a folder of excel sheets and organize a
% single chosen parameter by experiment in a single sheet. This will
% dramatically reduce the amount of time needed for data organization and
% analysis. For certain parameters, it will be necessary to have each cell (ID)
% separated, so there will be a variable in RunTime to select this option
% 

clear;
clc;

%Prompts user to select folder of interest, collects file names, then
%returns to HomeFolder

HomeFolder = cd;
DataFilePath = uigetdir;
cd(DataFilePath);
DataFiles = dir('*.xls');

cd(HomeFolder);

worksheetName = input('Type Excel Worksheet name of interest (case sensitive): ');
maxNumberTimePts = input('What is the maximum number of time points per each track?: ');

% worksheetName = "Ellipticity (prolate)"; Testing lines
% maxNumberTimePts = 40;

masterArray = [];
namesOfFiles = {};

skipStr = "Results sheet.xls";

for ii = 1:length(DataFiles)

    %OutputArray = NaN(maxNumberTimePts, 1);
    
    filename = strcat(DataFilePath,'/',DataFiles(ii).name);
    
    if ~(contains(filename,skipStr))
        
        %Read in data from the specified worksheet
        NumericData = readmatrix(filename, 'Sheet', worksheetName);
        TextHeaders = readcell(filename, 'Sheet', worksheetName);
        TextHeaders = TextHeaders (2,:);
        
        TimeIndex = find(strcmp(TextHeaders,'Time'));
        TrackIndex = find(strcmp(TextHeaders, 'TrackID'));
        
        AbridgedData = NumericData(:, [1 TimeIndex TrackIndex]);
        uniqueID = unique(AbridgedData(:,3));
        uniqueID = uniqueID';
        
        %Take the input sheet and split into columns for each unique track ID
        [~,~,X] = unique(AbridgedData(:,3));
        C = accumarray(X,1:size(AbridgedData,1),[],@(r){AbridgedData(r,:)});
        [nrows, ~] = cellfun(@size,C);
        
        AppendingArray = [];
        
        for j = 1:length(C)
            OutputArray = NaN(maxNumberTimePts, 1);
            currentCellMatrix = C{j};
            OutputArray(1:nrows(j),:) = currentCellMatrix(:,1);
            AppendingArray = [AppendingArray, OutputArray]; %#ok<*AGROW> %Warning is unnecessary
            
            
        end
        
        [appenRows, appenCols] = size(AppendingArray);
        SpacerArray = NaN(maxNumberTimePts+1,1);
        AppendingArray = [uniqueID; AppendingArray];
        AppendingArray = [AppendingArray, SpacerArray];
        
        
        nameOnTop = DataFiles(ii).name;
        [~, masterCol] = size(masterArray);
        namesOfFiles{masterCol+1} = nameOnTop; %#ok<*SAGROW> %Warning is unnecessary
        
        
        masterArray = [masterArray, AppendingArray];
    end
end
cd(DataFilePath);
newFileName = worksheetName + ' Results sheet.xls';

writecell(namesOfFiles, newFileName, 'Range', 'A1');
writematrix(masterArray, newFileName, 'Range', 'A2');
