function varargout = main(varargin)
% Written by Josh McGuckin (Associate Research Engineer,Spring/Summer 2021)
% 06/02/2022

% MAIN MATLAB code for main.fig
%      MAIN, by itself, creates a new MAIN or raises the existing
%      singleton*.
%
%      H = MAIN returns the handle to a new MAIN or the handle to
%      the existing singleton*.
%
%      MAIN('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in MAIN.M with the given input arguments.
%
%      MAIN('Property','Value',...) creates a new MAIN or raises
%      the existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before main_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to main_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help main

% Last Modified by GUIDE v2.5 19-Jul-2021 16:29:26

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @main_OpeningFcn, ...
                   'gui_OutputFcn',  @main_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT

% --- Executes just before main is made visible.
function main_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to main (see VARARGIN)

% Choose default command line output for main
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes main wait for user response (see UIRESUME)
% uiwait(handles.figure1);
set ( gcf, 'Color', [1 1 1] )
set (0,'DefaultFigureColor',[1 1 1])

axes(handles.globus_logo)
imshow('Globus logo.png');


% --- Outputs from this function are returned to the command line.
function varargout = main_OutputFcn(hObject, eventdata, handles)
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes when selected object changed in unitgroup.
function unitgroup_SelectionChangedFcn(hObject, eventdata, handles)
% hObject    handle to the selected object in unitgroup 
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
global facingoptotrak
if (hObject == handles.towards_optotrak)
    facingoptotrak = logical(1);
else
    facingoptotrak = logical(0);
end

% --------------------------------------------------------------------

% --- Executes on button press in run_program.
function run_program_Callback(hObject, eventdata, handles)
% hObject    handle to run_program (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
try
    % save specimen direction before erasing all global variables 
    global facingoptotrak
    if isempty(facingoptotrak)
        facingoptotrak = logical(1); % by default, the specimen is assumed to be facing the Optotrak system
    end
    specimendoesfaceoptotrak= facingoptotrak;
    clear global
    global facingoptotrak
    facingoptotrak = specimendoesfaceoptotrak;
    
    CloseProgram = 0;
    while CloseProgram == 0
        clear
        CloseProgram = 0; % remains 0

        %% UI for user to specify:
        %       test catagory order found in specified file 
        %       iteration distance between motion peaks
        prompt = {sprintf('%s\n%s\n\n%s\n%s','Before starting, ensure files are in alphanumeric order. Windows does not always do this.',...
            '(See the Six_DOF_ROM_Analyzer Documentation for more information.)',... 
            'Enter Test Catagory Names in Order Seperated by Commas:',... 
            '(Note: Acceptable names must start with a letter and may contain numbers but no special characters.)'),...
            sprintf('%s\n%s\n%s\n%s\n%s\n%s',...
            'Enter Graph Numbers Seperated by Commas for Manual Selection of Minimum and Maximum Values of 3rd Cycle:',...
            'Point Selection Instructions:','1) Zoom and move to find MAXIMUM peak location of 3rd cycle','2) Press Enter',...
            '3) Click and drag to form box that contains MAXIMUM peak point of 3rd cycle',...
            '4) Repeat steps 1-3 to now select the MINIMUM trough point of 3rd cycle')};
        dlgtitle = 'Input';
        dims = [1 105]; 

        global answer folderdir calfiledir facingoptotrak % saves the inputted test construct names list so that user doesn't need to re-enter them
        % if user restarts the program, use the previously inputted list of test catagory names 
        if isempty(answer)
            % UI for user to select folder containing the cal files
            folderdir = uigetdir(path,'Select Folder Containing cal.xls Files');
            calfiledir = dir(fullfile(folderdir,'*_cal.xls*'));
            answer = inputdlg(prompt,dlgtitle,dims);
        else
            answer = inputdlg(prompt,dlgtitle,dims,{answer{1},''});
        end
        
        %-- Load Bar --% 
        dlgwindow = uifigure;
        loadbarHandle = uiprogressdlg(dlgwindow,'Message','Checking Inputs...');
        %-- Load Bar --%
        answer1 = strrep(answer{1},' ','_'); % if there are underscores, replace with spaces to make them acceptable headers in table
        testCatList = strsplit(answer1,',');

        answer2 = str2num(strrep(answer{2},' ','')); % extract graph numbers to redo
        [manualSelNums,~] = sort(unique(answer2),'ascend'); % remove repetitions and put in ascending order
        if isempty(manualSelNums)
            manualSelNums = 0;
        end

        % test to see if valid variable name for xlswrite function
        for z = 1:length(testCatList)
            if testCatList{z}(1) == '_'
                testCatList{z}(1) = [];
            end
            if testCatList{z}(end) == '_'
                testCatList{z}(end) = [];
            end
            if ~isvarname(testCatList{z})
                errordlg(sprintf('%s \n%s\n%s','An inputted catagory name is not an acceptable MATLAB variable name.',...
                    'Acceptable names start with a letter and can have spaces, but no special characters.','Please Try Again.'))
                close(dlgwindow)
                return
            end
        end

        %check if # of test catogories inputted = # of xls cal files in folder
        % if not, # of user specified catogories is incorrect and elicit an error and stop script execution
        if ~(length(calfiledir)/3 == length(testCatList))
            errordlg(sprintf('%s \n%s %f %s \n%s','The number of test catagories specified is incorrect.','There are',...
                length(calfiledir)/3, 'test catagories.', 'Please Try Again.'))
            close(dlgwindow)
            return
        end
        %% 
        %-- Load Bar --%
        loadbarHandle.Value = 0.1; 
        loadbarHandle.Message = 'Analyzing .cal Files';
        %-- Load Bar --%

        % create subfolder in selected folder to store generated graphs
        newfolderdir = dir(fullfile(folderdir,'Relative Position Plots with Overlayed Min and Max Values*'));
        newFolderName = 'Relative Position Plots with Overlayed Min and Max Values';
        if isempty(newfolderdir)
            mkdir(folderdir, newFolderName); 
        end
        
        %% If Specimen is Facing Away from Optotrak System, Change Signage 
        if ~facingoptotrak
            FELBAR_signs = [1 -1 -1]; % corresponding signs for LB,FE,AR when calc relative distances
        else
            FELBAR_signs = [-1 1 -1]; % corresponding signs for LB,FE,AR when calc relative distances
        end
        
        %%
        FELBAR_index_names = {'LB (Rz)','FE (Ry)','AR (Rx)'};
        FELBAR_index = 2; % 1 = LB (Rz), 2 = FE (Ry), 3 = AR (Rx) --> starts on FE, then to LB, then to AR, then repeats
        relative_Index = 1; 

        matInsIndex = 1; % initialize matrix insert index
        graphNumIndex = 1;
        for i = 1:length(calfiledir)
            curFilePth = fullfile(folderdir, calfiledir(i).name); % index current file path
            % convert .xls to .txt file since dataset is too large
            txtFilePth = strrep(curFilePth, '.xls', '.txt'); % name of txt file

            %% DISCOVER WHAT TYPE OF FILE IS CURRENT '.xls' FILE & IMPORT DATA
            % OPTION 1: is it a true xls file?
            status = xlsfinfo(curFilePth);
            if ~isempty(status)
                [data,colheaders,~] = xlsread(curFilePth);
                data = data(2:end,:);
            % OPTION 2: it is truly a txt file with an .xls extension 
            else 
                if ~strcmp(curFilePth,txtFilePth)
                    movefile(curFilePth,txtFilePth) % overwrite filename to .txt
                end
                txtFileStruct = importdata(txtFilePth);
                % import the new txt file data
                data = txtFileStruct.data;
                % if user opened the fake xls file, edited it, and saved,
                % then the colheaders will be in quotes - which requires
                % different indexing
                if length(txtFileStruct.textdata) == 1
                    colheaders = txtFileStruct.textdata{:}; % seperate colheaders
                    colheaders = strsplit(colheaders,'\t');
                else
                    colheaders = txtFileStruct.textdata; % edited fake xls file has its colheaders seperated already
                    colheaders = strrep(colheaders,'"',''); % remove the quotes 
                end
            end

            %% REMOVE EMPTY COLUMNS WITH RELATIVE COLUMN HEADERS
            isnancols = find(all(isnan(data),1));
            data(:,isnancols) = []; % if any columns are completely NaN, remove them
            colheaders(isnancols) = []; % if any columns are completely NaN, remove them
            %% REMOVE COLUMNS FROM UNUSED MARKERS WHICH ACCIDENTALLY BECAME IN VIEW 
            % such markers will not be in view at the very beginning of data
            % collection, but may appear transiently throughout data acquisition  
            data(:, isnan(data(1,:)))  = [];
            colheaders(isnan(data(1,:))) = [];
            %% FIND NUMBER OF RELATIVES/COLUMNS
            [~, numcols] = size(data);
            colheaders = colheaders(1:numcols); % removes the extra relative column headers 

            %% PLOTTING RELATIVE DISTANCES, SELECTING PEAKS, AND PLOTTING 
            if i == 1 % set up number of relatives (i.e., summary files)
                numrelatives = (numcols - 1)/6;
                summarydatacells = cell(1, numrelatives);
                rowheaders = {'Flexion','Extension','Right Bending','Left Bending','Left Rotation','Right Rotation','Flexion - Extension'... 
                    'Lateral Bending','Axial Rotation'}';
            end

            for j = 0:(numrelatives-1)
                distances = data(:,(1+FELBAR_index+6*j));
                relDistances = FELBAR_signs(FELBAR_index)*(distances - distances(1));
                % collect the names of the relatives being captured
                if i == 1
                    relativeName = colheaders{1+FELBAR_index+6*j};
                    relativeName = strsplit(relativeName,' deg');
                    relativeNames{j+1} = relativeName{1};
                end

                %% PLOT AND SAVE RELATIVE DISTANCES & MIN AND MAX VALUES WITHOUT DISPLAYING THEM
                % if user wants, UI should open to select peaks manually. else, it's automated
                if any(graphNumIndex == manualSelNums)
                    % PLOT GRAPH AND DISPLAY IT
                    curFig = figure('visible','on');
                    plot(relDistances)
                    cur_axis = axis; % get current axis to then reset it after zooming in
                    plotTitle = sprintf('Graph %d) %s - %s %s',graphNumIndex, testCatList{ceil(i/3)}, FELBAR_index_names{FELBAR_index},...
                        relativeNames{relative_Index});
                    plotTitle = strrep(plotTitle, '_', ' '); % remove underscores from title as they make proceeding letters subscripts
                    title(plotTitle)
                    xlabel('Data Point')
                    ylabel('Relative Angle (°)')

                    % ALLOW USER TO ZOOM
                    buttonwait = 0;
                    while ~buttonwait
                        buttonwait = waitforbuttonpress;
                        if ~strcmp(get(gcf,'CurrentKey'),'return')
                            buttonwait = 0;
                        end
                    end
                    rect = getrect(curFig); % [xmin ymin width height]
                    % FIND MAX VALUE INSIDE RECTANGLE THAT ISN'T HIGHER THAN RECTANGLE
                    rectDomain = 1:length(relDistances);
                    rectDomain = rectDomain( rectDomain >= floor(rect(1)) & rectDomain <= ceil((rect(1)+rect(3))) );
                    relDistLocsInROI = rectDomain(relDistances(rectDomain) <= ( rect(2)+rect(4) ) & relDistances(rectDomain) >= ( rect(2) ));
                    [maxDeg, I] = max(relDistances(relDistLocsInROI));
                    maxDegloc = relDistLocsInROI(I);
                    % OVERLAY SELECTED MAX POINT
                    hold on
                    plot(maxDegloc,maxDeg,'*r','MarkerSize',8)
                    hold off
                    % ALLOW USER TO ZOOM
                    buttonwait = 0;
                    while ~buttonwait
                        buttonwait = waitforbuttonpress;
                        if ~strcmp(get(gcf,'CurrentKey'),'return')
                            buttonwait = 0;
                        end
                    end
                    rect = getrect(curFig); % [xmin ymin width height]
                    % FIND MIN VALUE INSIDE RECTANGLE THAT ISN'T LOWER THAN RECTANGLE 
                    rectDomain = 1:length(relDistances);
                    rectDomain = rectDomain( rectDomain >= floor(rect(1)) & rectDomain <= ceil((rect(1)+rect(3))) );
                    relDistLocsInROI = rectDomain(relDistances(rectDomain) <= ( rect(2)+rect(4) ) & relDistances(rectDomain) >= ( rect(2) ));
                    [minDeg, I] = min(relDistances(relDistLocsInROI));
                    minDegloc = relDistLocsInROI(I);
                    % OVERLAY SELECTED MIN POINT
                    hold on
                    plot(minDegloc,minDeg,'*r','MarkerSize',8)
                    hold off
                    axis(cur_axis) % reset axis before saving image
                else
                    % FIND 3 MAIN PEAKS
                    [smoothpks,smoothlocs] = findpeaks(relDistances,'MinPeakDistance',50); % smoother data
                    [pks,locs,widths,proms] = findpeaks(smoothpks,'Annotate','extents');
                    approxAreas = widths.*proms;% approximate areas under cycles 
                    [~,r] = sort(approxAreas,'ascend'); % use ascend (instead of descend) to index the last 3 elements which are later data pts 
                    locsOfWidths = smoothlocs(locs(r(end-2:end))); % locs containing 3 peaks of cycles with the 3 largest widths
                    maxDegloc = max(locsOfWidths); % finds last (3rd) peak of the 3 widest cycles
                    maxDeg = relDistances(maxDegloc); 
                    % FIND 3 MAIN TROUGHS
                    [smoothtroughs,smoothlocs] = findpeaks(-relDistances,'MinPeakDistance',50); % smoother data
                    [troughs,locs,widths,proms] = findpeaks(smoothtroughs,'Annotate','extents');
                    approxAreas = widths.*proms;
                    [~,r] = sort(approxAreas,'ascend');
                    locsOfWidths = smoothlocs(locs(r(end-2:end)));
                    minDegloc = max(locsOfWidths);
                    minDeg = relDistances(minDegloc);      
                    % PLOT GRAPH AND OVERLAY MIN & MAX VALUES WITHOUT DISPLAYING IT
                    curFig = figure('visible','off');
                    plot(relDistances)
                    plotTitle = sprintf('Graph %d) %s - %s %s',graphNumIndex, testCatList{ceil(i/3)}, FELBAR_index_names{FELBAR_index},...
                        relativeNames{relative_Index});
                    plotTitle = strrep(plotTitle, '_', ' '); % remove underscores from title as they make proceeding letters subscripts
                    title(plotTitle)
                    xlabel('Data Point')
                    ylabel('Relative Angle (°)')
                    hold on
                    plot([minDegloc maxDegloc],[minDeg maxDeg],'*r','MarkerSize',8)
                    hold off
                end

                %% SAVE AND CLOSE PLOT
                saveas(curFig,fullfile(folderdir,newFolderName, [plotTitle '.jpg']),'jpg'); % saves to new plots subfolder
                close(curFig)

                %% SAVE MAX & MIN VALUES INTO SUMMARY DATA CELLS
                if i == 1
                    summarydatacells{1,j+1} = zeros(6,length(testCatList)); % pre-define sizes of each summary sheet (saves time)
                end
                % store min and max values in the summary data cells that will become summary data sheets
                summarydatacells{1,j+1}(matInsIndex) = maxDeg;
                summarydatacells{1,j+1}(matInsIndex+1) = minDeg;

                % update relatives index
                relative_Index = relative_Index+1;
                if relative_Index > numrelatives
                    relative_Index = 1;
                end
                % update graph number index 
                graphNumIndex = graphNumIndex+1;
            end
            % update FELBAR index 
            FELBAR_index = FELBAR_index +2;
            if FELBAR_index > 3
                FELBAR_index = FELBAR_index-3;
            end
            % update matrix insert index
            matInsIndex = matInsIndex +2; 

            % if applicable, rename txt file back to xls extension
            if isempty(status) && ~strcmp(curFilePth, txtFilePth)
                xlsFilePth = strrep(curFilePth, '.txt','.xls');
                movefile(txtFilePth,xlsFilePth) % overwrite filename
            end 

            %-- Load Bar --%
            loadbarHandle.Value = 0.1 + (i/length(calfiledir))*0.8;
            %-- Load Bar --%
        end
        
        % add the last 3 rows of the summary sheets (i.e., flexion-extention, lateral bending, and axial rotation)
        for k = 1:length(summarydatacells)
            summarydatacells{1,k}(end+1,:) = summarydatacells{1,k}(1,:) - summarydatacells{1,k}(2,:);
            summarydatacells{1,k}(end+1,:) = summarydatacells{1,k}(3,:) - summarydatacells{1,k}(4,:);
            summarydatacells{1,k}(end+1,:) = summarydatacells{1,k}(5,:) - summarydatacells{1,k}(6,:);
        end
        
        %% Save Specimen Summary Sheets into an Excel File
        % excel files: summarydata cells has the data for each summary file 
        % colheaders contains column headers for ALL summary files
        % rowheaders contains the row headers for ALL summary files 

        %-- Load Bar --%
        loadbarHandle.Value = 0.9; 
        loadbarHandle.Message = 'Making Excel Summary Sheets';
        %-- Load Bar --%
        
        xlsfullfile = fullfile(folderdir,'Specimen Summary.xls');
        % create Excel File with Summary sheets
        for z = 1:length(summarydatacells)
            % create sheet name that specifies its corresponding relatives
            sheetName = sprintf('Summary %s',relativeNames{z});
            % write Excel file
            T = num2cell(summarydatacells{1,z});
            T = [testCatList;T];
            T = [['Rows';rowheaders] T];
            xlsT = cell2table(T(2:end,:));
            xlsT.Properties.VariableNames = T(1,:);
            writetable(xlsT,xlsfullfile,'Sheet',sheetName);
        end

        %% notify user that script has executed completely. give user option to restart program
        close(loadbarHandle)
        quest = sprintf('%s \n\n%s','The summary Excel file and subfolder containing generated plots are now located in the selected folder.',...
            'Do you want to restart and manually select points for specific graphs?');
        questitle = 'Analysis Complete!';
        btn1 = 'Done';
        btn2 = 'Restart';
        answer2 = uiconfirm(dlgwindow,quest,questitle,'Options',{btn1,btn2},'Icon','Question');
        if strcmp(answer2,btn1)
            CloseProgram = 1; % Closes Program
        end
        close(dlgwindow)
    end
catch
        % if applicable, rename txt file back to xls extension before program crashes
        calfiledir = dir(fullfile(txtFilePth));
        if ~isempty(calfiledir)
            movefile(txtFilePth,curFilePth) % overwrite file from txt back to xls 
        end 
        errordlg('Program Has Failed. Please Try Again.')
        close(dlgwindow)
end
