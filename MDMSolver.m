clc;clear;close all;
PathAdd=input('PLEASE ENTER THE PATH FOR TEST CASE FOLDER:','s');
TestCase=input('PLEASE ENTER THE NAME OF THE FOLDER WITH TEST CASES: ','s');
addpath(PathAdd);
% Specify the folder where the files live.
MyFolder=[PathAdd,'\',TestCase];
% Check to make sure that folder actually exists.  Warn user if it doesn't.
while ~isdir(MyFolder)
    disp(' ');disp('ERROR:PLEASE MAKE SURE YOU HAVE ENTERED A CORRECT PATH AND FOLDER NAME');disp(' ');
    PathAdd=input('PATH FOR TEST CASE FOLDER:','s');
    TestCase=input('NAME OF TEST CASE FOLDER: ','s');
    addpath(PathAdd);
    MyFolder =[PathAdd,'\',TestCase];
end
%for result folder
AnsFolder=input('GIVE A NAME FOR RESULT FOLDER: ','s');
MyTest=[PathAdd,'\',AnsFolder];
TempVal_1=1;
while isdir(MyTest)
    disp('  ');
    if (TempVal_1)
        disp('WARNING: THIS RESULT FOLDER ALREADY EXISTS');
    end
    temp_var=input('DO YOU WANT TO CONTINUE TO OVERWRITE THE RESULTS?(Y/N) :','s');
    if (strcmp(temp_var,'Y') || strcmp(temp_var,'y'))
        warning off MATLAB:MKDIR:DirectoryExists
        break;
    elseif (strcmp(temp_var,'N') || strcmp(temp_var,'n'))
        AnsFolder=input('GIVE A NAME FOR RESULT FOLDER: ','s');TempVal_1=1;
    else
        disp('PLEASE ENTER A VALID OPTION');TempVal_1=0;
    end
    MyTest=[PathAdd,'\',AnsFolder];
end
mkdir([PathAdd,'\',AnsFolder]);
% Get a list of all files in the folder with the desired file name pattern.
FilePattern = fullfile(MyFolder, '*.xlsx'); % Change to whatever pattern you need.
TheFiles = dir(FilePattern);
warning('off','MATLAB:xlswrite:AddSheet');
%search for each excel file in the folder

for files = 1:length(TheFiles)
    baseFileName = TheFiles(files).name;
    flname = fullfile(MyFolder, baseFileName);
    %extracting excel file names in the folder
    temp_vals_1=strfind(flname,'\');temp_vals_2=strfind(flname,'.');
    file=flname(temp_vals_1(end)+1:temp_vals_2(end)-1);
    flname2=flname(1:temp_vals_1(end-1)-1);
    %escape the temporary file
    if (strcmp(file(1:2),'~$'))
        continue;
    end
    disp('   ');
    disp('-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-');
    disp(['TEST CASE FILE: ',file]);
    disp('-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-');disp(' ');
    
    disp('Reading The File');
    [values,txt,raw] = xlsread(flname);%for reading the excel file
    File=[flname2,'\',AnsFolder,'\',file,'_Solution.xlsx'];
    xlswrite(File,txt,'Problem MDM');
    OutputForUserInput=txt;
    NumberOfDSM=zeros(1,1);
    if (isnan(raw{1,1}) || ischar(raw{1,1}))
        msgbox('Incorrect File Format',file,'error');
        disp('File Format Is Not Correct,Please Check The Format And Try Again');
        continue;
    end
    %Extracting the number of dmms
    MatrixSize=size(txt,1)-1;TempSum=0;NumberOfDSM(1)=raw{1,1};
    TempVar=0;
    for DMM=1:raw{1,1}
        if (isnan(raw{DMM+1,1}) || ischar(raw{DMM+1,1}))
            msgbox('Incorrect File Format',file,'error');
            disp('File Format Is Not Correct,Please Check The Format And Try Again');
            TempVar=1;break;
        else
            NumberOfDSM(DMM+1)=raw{DMM+1,1};
            TempSum=TempSum+raw{DMM+1,1};
        end
    end
    if (TempVar)
        continue;
    end
    if (MatrixSize==TempSum)
        disp('File Format Is Correct');
    else
        msgbox('Incorrect File Format',file,'error');
        disp('File Format Is Not Correct,Please Check The Format And Try Again');
        continue;
    end
    %Extracting the actual matrix
    MDMMatrix=zeros(MatrixSize);
    %replacing blank space with 1 (including the diagonal) and X with
    for Row=2:MatrixSize+1
        for Column=2:MatrixSize+1
            if (strcmp(txt{Row,Column},'X') || strcmp(txt(Row,Column),'x'))
                MDMMatrix(Row-1,Column-1)=1;
            else
                MDMMatrix(Row-1,Column-1)=0;
            end
        end
    end
    
    %matix multipilcation
    disp('Calculating For The Higher DSM(s)');
    for DSM=1:NumberOfDSM(1)-1
        From1=sum(NumberOfDSM(2:DSM))+1; % from row (for horizontal (above diagonal) matrix
        To1=From1+NumberOfDSM(DSM+1)-1;    % upto row (for horizontal (above diagonal) matrix
        From2=To1+1;MaxNumberOfX=0;MaxNumberOfXMatrix=zeros(NumberOfDSM(DSM+1));
        for DMM=1:NumberOfDSM(1)-DSM
            NumberOfX=0;To2=From2+NumberOfDSM(DSM+DMM+1)-1;
            TempMatrixHorizontal=MDMMatrix(From1:To1,From2:To2);
            TempMatrixVertical=MDMMatrix(From2:To2,From1:To1);
            TempDSM=TempMatrixHorizontal*TempMatrixVertical;
            From2=To2+1;
            %counting the number of X
            for Row=1:NumberOfDSM(DSM+1)
                for Col=1:NumberOfDSM(DSM+1)
                    if (Row~=Col && TempDSM(Row,Col)~=0)
                        NumberOfX=NumberOfX+1;
                    end
                end
            end
            %chossing max X matrix
            if (NumberOfX>MaxNumberOfX)
                MaxNumberOfX=NumberOfX;
                MaxNumberOfXMatrix=TempDSM;
            end
        end
        for Row=1:NumberOfDSM(DSM+1)
            for Col=1:NumberOfDSM(DSM+1)
                if (Row~=Col && MaxNumberOfXMatrix(Row,Col)~=0)
                    OutputForUserInput{From1+Row,From1+Col}='X';
                    MDMMatrix(From1+Row-1,From1+Col-1)=1;
                else
                    OutputForUserInput{From1+Row,From1+Col}='';
                    MDMMatrix(From1+Row-1,From1+Col-1)=0;
                end
            end
        end
    end
    disp('Obtained All The Higher DSM(s)');
    disp('Writing in Excel File For The User Input');
    Excel = actxserver('Excel.Application');
    WB=invoke(Excel.Workbooks,'Open',File);
    xlswrite1(File,OutputForUserInput,'DMM Multiplication');
    for DSM=1:NumberOfDSM(1)
        From1=sum(NumberOfDSM(2:DSM))+2; % from row (for horizontal (above diagonal) matrix
        To1=From1+NumberOfDSM(DSM+1)-1;    % upto row (for horizontal (above diagonal) matrix
        wsheet=get(Excel,'ActiveSheet');
        range=get(wsheet,'Range',[ColumnName(From1),num2str(From1),':',ColumnName(To1),num2str(To1)]);
        borders=get(range,'Borders');set(borders,'ColorIndex',3);set(borders,'LineStyle',6);
    end
    for Row=1:MatrixSize
        WB.Worksheets.Item('DMM Multiplication').Range([ColumnName(Row+1),num2str(Row+1),':',ColumnName(Row+1),num2str(Row+1)]).Interior.ColorIndex =6;
    end
    try
        Excel.ActiveWorkbook.Worksheets.Item('Sheet1').Delete;
    catch
        %Do Nothing
    end
    WB.Save();% Save Workbook
    WB.Close();% Close Workbook
    disp('Opening The Excel File For User Input');
    winopen(File);
    %now wait for the user input
    disp('Excel File Opened For The User Input');
    disp('Please Enter All The Diagonal Values In "DMM Multiplication" Sheet');disp(' ');
    disp('Step1:After Entering All The Diagonal Values,Please Save And Close The Excel File.')
    disp('Step2:Click AnyWhere On The Command Window And Press Any Key If You Are Done With Step1');
    pause;disp(' ');
    disp('Assuming You Have Done Correctly,Procceding For Block Formation');
    disp('Block Formation In Progress');
    %block formation in each DSM
    
    values=xlsread(File,'DMM Multiplication');%for reading the excel file
    for DSM=1:MatrixSize
        MDMMatrix(DSM,DSM)=values(DSM,DSM);
    end
    
    AllBlockAddress=cell(NumberOfDSM(1)-1,1);
    OutputForUserReworkInput=cell(MatrixSize+1);IsValidCase=1;WantToContinue=1;
    for DSM=NumberOfDSM(1):-1:1
        From=sum(NumberOfDSM(2:DSM))+1; % from row (for horizontal (above diagonal) matrix
        To=From+NumberOfDSM(DSM+1)-1;    % upto row (for horizontal (above diagonal) matrix
        
        if (DSM>1)
            TempMatrix=MDMMatrix(From:To,From:To);
            [BlockAddress,SequenceNew,IsValid,Continue]=Clustering(TempMatrix,file);
            if (~Continue)
                WantToContinue=0;break;
            end
            TempSequance=From:To;
            TempSequance=TempSequance(:,SequenceNew);
            MDMMatrix=MDMMatrix(:,[1:From-1,TempSequance]);
            MDMMatrix=MDMMatrix([1:From-1,TempSequance],:);
            if (~IsValid)
                IsValidCase=0;
                for Row=1:MatrixSize
                    for Col=1:MatrixSize
                        if (Row~=Col && MDMMatrix(Row,Col)~=0)
                            OutputForUserReworkInput{Row+1,Col+1}='X';
                        elseif (Row~=Col)
                            OutputForUserReworkInput{Row+1,Col+1}='';
                        else
                            OutputForUserReworkInput{Row+1,Col+1}=MDMMatrix(Row,Col);
                        end
                    end
                end
                for Row=1:To
                    if (Row<From)
                        OutputForUserReworkInput{Row+1,1}=txt{Row+1,1};
                        OutputForUserReworkInput{1,Row+1}=txt{Row+1,1};
                    else
                        OutputForUserReworkInput{Row+1,1}=txt{TempSequance(Row-From+1)+1,1};
                        OutputForUserReworkInput{1,Row+1}=txt{TempSequance(Row-From+1)+1,1};
                    end
                end
                OutputForUserReworkInput{MatrixSize+4,1}='RESULT:';
                OutputForUserReworkInput{MatrixSize+4,2}='Invalid MDM (All X is Below the Diagonal in DSM Marked Red)';
                WB=invoke(Excel.Workbooks,'Open',File);
                
                xlswrite1(File,OutputForUserReworkInput,'Block Formation');
                WB.Worksheets.Item('Block Formation').Range([ColumnName(From+1),num2str(From+1),':',ColumnName(To+1),num2str(To+1)]).Interior.ColorIndex =3;
                WB.Worksheets.Item('Block Formation').Range(['A',num2str(MatrixSize+4)]).Interior.ColorIndex =3;
               
                WB.Save();% Save Workbook
                WB.Close();% Close Workbook
                Excel.Quit();% Quit Excel
                break;
            end
            AllBlockAddress{DSM-1}=BlockAddress;
            for Row=From:To
                OutputForUserReworkInput{Row+1,1}=txt{TempSequance(Row-From+1)+1,1};
                OutputForUserReworkInput{1,Row+1}=txt{TempSequance(Row-From+1)+1,1};
            end
        else
            for Row=From:To
                OutputForUserReworkInput{Row+1,1}=txt{Row+1,1};
                OutputForUserReworkInput{1,Row+1}=txt{Row+1,1};
            end
        end
    end
    if (~IsValidCase)
        msgbox('Invalid MDM Mtrix,Terminating The Process For The MDM',file,'error');
        disp('Invalid MDM Mtrix,Process Terminated, Please Check the MDM And Run Again');
        continue;
    end
    if (~WantToContinue)
        disp('You Choose Not To Proceed Further,Terminating the Process');
        OutputForUserReworkInput{1}='Process Terminated By The User As Forming The Blocks Would Have Taken Significant Time';
        disp('Process Terminated');
        WB=invoke(Excel.Workbooks,'Open',File);
        xlswrite1(File,OutputForUserReworkInput,'Block Formation');
        WB.Worksheets.Item('Block Formation').Range('A1:I1').Interior.ColorIndex =3;
   
        WB.Save();% Save Workbook
        WB.Close();% Close Workbook
        Excel.Quit();% Quit Excel
        continue;
    end
    disp('Done With The Block Formation');
    disp('Writing In Excel File For The User Input');
    for Row=1:MatrixSize
        for Col=1:MatrixSize
            if (Row~=Col && MDMMatrix(Row,Col)~=0)
                OutputForUserReworkInput{Row+1,Col+1}='X';
            elseif (Row~=Col)
                OutputForUserReworkInput{Row+1,Col+1}='';
            else
                OutputForUserReworkInput{Row+1,Col+1}=MDMMatrix(Row,Col);
            end
        end
    end
    WB=invoke(Excel.Workbooks,'Open',File);
    xlswrite1(File,OutputForUserReworkInput,'Block Formation');
    %forming the Border for dsm
    for DSM=1:NumberOfDSM(1)
        From1=sum(NumberOfDSM(2:DSM))+2; % from row (for horizontal (above diagonal) matrix
        To1=From1+NumberOfDSM(DSM+1)-1;    % upto row (for horizontal (above diagonal) matrix
        wsheet=get(Excel,'ActiveSheet');
        range=get(wsheet,'Range',[ColumnName(From1),num2str(From1),':',ColumnName(To1),num2str(To1)]);
        borders=get(range,'Borders');set(borders,'ColorIndex',3);set(borders,'LineStyle',6);
    end
    %highlighting the diagonals
    for Row=1:MatrixSize
        WB.Worksheets.Item('Block Formation').Range([ColumnName(Row+1),num2str(Row+1),':',ColumnName(Row+1),num2str(Row+1)]).Interior.ColorIndex =6;
    end
    %highlighting the blocks in each dsm
    for DSM=NumberOfDSM(1):-1:2
        From=sum(NumberOfDSM(2:DSM))+1; % from row (for horizontal (above diagonal) matrix
        for Block=1:size(AllBlockAddress{DSM-1},1)
            ColFrom=From+AllBlockAddress{DSM-1}(Block,1);
            ColTo=From+AllBlockAddress{DSM-1}(Block,2);
            WB.Worksheets.Item('Block Formation').Range([ColumnName(ColFrom),num2str(ColFrom),':',ColumnName(ColTo),num2str(ColTo)]).Interior.ColorIndex =7;
        end
    end
    WB.Save();% Save Workbook
    WB.Close();% Close Workbook
    disp('Opening The Excel File For User Input');
    winopen(File);
    %now wait for the user input
    disp('Excel file opened for user input');
    disp('Please Enter the Rework Probability In "Block Formation" Sheet');disp(' ');
    disp('Step1:After Entering All the Rework Probability,Please Save And Close The Excel File.')
    disp('Step2:Click AnyWhere On The Command Window And Press Any Key If You Are Done With Step1');
    pause;disp(' ');
    disp('Assuming You Have Done Correctly,Procceding For Markov Chain');
    disp('Markov Chain In Progress');
    [values,txt,raw]=xlsread(File,'Block Formation');
    WB=invoke(Excel.Workbooks,'Open',File);
    for DSM=NumberOfDSM(1):-1:2
        
        From=sum(NumberOfDSM(2:DSM))+1; % from row (for horizontal (above diagonal) matrix
        for Block=1:size(AllBlockAddress{DSM-1},1)
            ColFrom=From+AllBlockAddress{DSM-1}(Block,1);
            ColTo=From+AllBlockAddress{DSM-1}(Block,2);
            TempName=cell(1,ColFrom-ColTo+1);
            for i=ColFrom:ColTo
                TempName{1,i-ColFrom+2}=txt{1,i};
            end
            Matrix=values(ColFrom-1:ColTo-1,ColFrom-1:ColTo-1);
            MatrixSize=size(Matrix,1);% TempArraySize=N, where TempArray is NXN matrix
            PermutationElements=1:MatrixSize;% permutation size
            PermutationList=perms(PermutationElements);% list of all permutations
            
            ListStore=cell(1,1);
            ListStore{1,1}='NO.'; ListStore{1,2}='SEQUENCE CODE';ListStore{1,3}='SEQUENCE NAME'; ListStore{1,4}='BLOCK DUARATION';
            % Max_Index=1;Min_Index=1;
            MinValue=Inf;MaxValue=-Inf;
            
            for PermutationNumber=1:size(PermutationList,1)
                TempMatrix=Matrix;OutputMatrix=cell(1,1);
                TempSequance=PermutationList(PermutationNumber,:);SequenceNameInWords='';SequanceNameInDigits='';
                for j=1:MatrixSize
                    SequenceNameInWords=strcat(SequenceNameInWords,TempName{1,PermutationList(PermutationNumber,j)+1});
                    if (j<MatrixSize)
                        SequenceNameInWords=strcat(SequenceNameInWords,'-');
                    end
                    SequanceNameInDigits=strcat(SequanceNameInDigits,num2str(PermutationList(PermutationNumber,j)));
                    OutputMatrix{1,j+1}=TempName{1,PermutationList(PermutationNumber,j)+1};
                    OutputMatrix{j+1,1}=TempName{1,PermutationList(PermutationNumber,j)+1};
                end
                TempMatrix=TempMatrix(:,TempSequance);
                TempMatrix=TempMatrix(TempSequance,:);
                
                %for writing raw matrix
                for Row=1:MatrixSize
                    for Col=1:MatrixSize
                        OutputMatrix{Row+1,Col+1}=TempMatrix(Row,Col);
                    end
                end
                
                TempMatrix(isnan(TempMatrix))=0;
                DiagonalVal=diag(TempMatrix);
                TransTempArray=-transpose(TempMatrix);
                TransTempArray(logical(eye(size(TransTempArray))))=1;
                %for writing second matrix
                for j=1:MatrixSize
                    for k=1:MatrixSize
                        OutputMatrix{j+MatrixSize+3,k+1}=TransTempArray(j,k);
                    end
                end
                
                TransTempArray_2=[TransTempArray,DiagonalVal];
                %solving for values using gaussian transformation
                %column wise
                for j=1:MatrixSize-1
                    %row wise
                    for k=j+1:MatrixSize
                        TransTempArray_2(k,:)=round(TransTempArray_2(k,:)-TransTempArray_2(j,:)/TransTempArray_2(j,j)*TransTempArray_2(k,j),2);
                    end
                end
                %cell storage
                for j=1:MatrixSize
                    for k=1:MatrixSize
                        OutputMatrix{j+2*MatrixSize+5,k+1}=TransTempArray_2(j,k);
                    end
                end
                TempVal=0;
                OutputMatrix{round(MatrixSize/2)+MatrixSize+3,MatrixSize+4}='=';
                OutputMatrix{round(MatrixSize/2)+2*MatrixSize+5,MatrixSize+4}='=';
                for j=1:MatrixSize
                    TempVal_1=TransTempArray_2(j,MatrixSize+1)/TransTempArray_2(j,j);
                    TempVal_1=round(TempVal_1,2);
                    TempVal=TempVal+TempVal_1;
                    %write second matrix
                    OutputMatrix{j+MatrixSize+3,MatrixSize+3}=['r',TempName{1,PermutationList(PermutationNumber,j)+1}];
                    OutputMatrix{j+MatrixSize+3,MatrixSize+5}=DiagonalVal(j);
                    %write gaussian eliminated matrix
                    OutputMatrix{j+2*MatrixSize+5,MatrixSize+3}=['r',TempName{1,PermutationList(PermutationNumber,j)+1}];
                    OutputMatrix{j+2*MatrixSize+5,MatrixSize+5}=TransTempArray_2(j,MatrixSize+1);
                    %write the answers
                    OutputMatrix{j+3*MatrixSize+7,MatrixSize+3}=['r',TempName{1,PermutationList(PermutationNumber,j)+1}];
                    OutputMatrix{j+3*MatrixSize+7,MatrixSize+4}='=';
                    OutputMatrix{j+3*MatrixSize+7,MatrixSize+5}=TempVal_1;
                end
                %write the final answer
                OutputMatrix{4*MatrixSize+8,MatrixSize+3}='TPD';
                OutputMatrix{4*MatrixSize+8,MatrixSize+4}='=';
                OutputMatrix{4*MatrixSize+8,MatrixSize+5}=TempVal;
                %for iteration duration
                ListStore{PermutationNumber+1,1}=PermutationNumber;ListStore{PermutationNumber+1,2}=SequanceNameInDigits;
                ListStore{PermutationNumber+1,3}=SequenceNameInWords;ListStore{PermutationNumber+1,4}=TempVal;
                % find max
                if (MinValue>TempVal)
                    MinValue=TempVal;
                end
                %find min
                if (MaxValue<TempVal)
                    MaxValue=TempVal;
                end 
                %write in excel sheet
                SheetName=['Trial-',SequanceNameInDigits,' (',SequenceNameInWords,')'];
                xlswrite1(File,OutputMatrix,SheetName);
                
                k=2;k1=2;
                for j=1:3*MatrixSize+1
                    if (j==3*MatrixSize+1)
                        TempRange=[ColumnName(MatrixSize+3),num2str(k+MatrixSize),':',ColumnName(MatrixSize+5),num2str(k+MatrixSize)];
                        WB.Worksheets.Item(SheetName).Range(TempRange).Interior.ColorIndex=4;
                    else
                        TempRange=[ColumnName(k1),num2str(k)];
                        WB.Worksheets.Item(SheetName).Range(TempRange).Interior.ColorIndex=6;% Set the color of cell "A1" of Sheet 1 to RED
                    end
                    if (rem(j,MatrixSize)==0)
                        k=k+3;k1=2;
                    else
                        k=k+1;k1=k1+1;
                    end
                end
            end
            %Done for all permutations,storing all min and maxes
            MaxIndexList=zeros(1);
            MinIndexList=zeros(1);
            MaxIndex=1;MinIndex=1;
            for i=1:size(PermutationList,1)
                if (MaxValue==ListStore{i+1,4})
                    MaxIndexList(MaxIndex)=i;
                    MaxIndex=MaxIndex+1;
                end
                if (MinValue==ListStore{i+1,4})
                    MinIndexList(MinIndex)=i;
                    MinIndex=MinIndex+1;
                end
            end
            SheetName=['IterationDuration(DSM',num2str(DSM),'Block',num2str(Block),')'];
            xlswrite1(File,ListStore,SheetName);
            %max row coloring
            for i=1:size(MaxIndexList,2)
                TempRange=['A',num2str(MaxIndexList(i)+1),':','D',num2str(MaxIndexList(i)+1)];
                WB.Worksheets.Item(SheetName).Range(TempRange).Interior.ColorIndex =4;
            end
            % min row coloring
            for i=1:size(MinIndexList,2)
                TempRange=['A',num2str(MinIndexList(i)+1),':','D',num2str(MinIndexList(i)+1)];
                WB.Worksheets.Item(SheetName).Range(TempRange).Interior.ColorIndex =6;
            end
        end
    end
    disp('Done With The Markov Chain');disp('Solved the MDM');
    WB.Save();WB.Close();% Save & Close Workbook
    Excel.Quit();
end
disp('*********************Done for All the MDM.Enojoy!!!*********************');