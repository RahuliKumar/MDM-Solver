function [BlockAddress,SequenceNew,IsValidCase,Continue]=Clustering(Matrix,FileName)
[OrderedMatrix,OrignalNewSequence,From,To]=RemoveBlankRowColumn(Matrix);
if (size(OrderedMatrix,1)>8)
    button = questdlg('Block Formation Will Take Time.Want To Continue?', FileName, 'Yes', 'No', 'Yes');
    if (strcmp(button,'Yes'))
        Continue=1;
        [BlockAddress,SequenceNew,IsValidCase]=BruteforceClustering(OrderedMatrix,OrignalNewSequence,From,To);
        BlockAddress=BlockAddress+From-1;
        
    else
        Continue=0;BlockAddress=0;SequenceNew=0;IsValidCase=1;
    end
else
    Continue=1;
    [BlockAddress,SequenceNew,IsValidCase]=BruteforceClustering(OrderedMatrix,OrignalNewSequence,From,To);
    BlockAddress=BlockAddress+From-1;
end
end

function [OrderedMatrix,OrignalNewSequence,From,To]=RemoveBlankRowColumn(UnorderedMatrix)
Loop=1;TempFrom=1;TempTo=size(UnorderedMatrix,1);
TruncatedOldSequance=1:TempTo;TempOrignalNewSequence=1:TempTo;
while (Loop)
    UnarrangedMatrixSize=size(UnorderedMatrix,1);Loop=0;
    if (UnarrangedMatrixSize>2)
        Temp_Seq=1:UnarrangedMatrixSize;
        ZeroRows=0;ZeroColumns=0;
        for Row=1:UnarrangedMatrixSize
            AllZero=1;
            for Col=1:UnarrangedMatrixSize
                if (Row~=Col && UnorderedMatrix(Row,Col)~=0)
                    AllZero=0;break;
                end
            end
            if (AllZero)
                ZeroRows=ZeroRows+1;
                Temp_Seq(Temp_Seq==Row)=[];
                Temp_Seq=[Row,Temp_Seq];
            end
        end
        if (ZeroRows>0)
            Loop=1;
            UnorderedMatrix=UnorderedMatrix(:,Temp_Seq);
            UnorderedMatrix=UnorderedMatrix(Temp_Seq,:);
%             disp('Row');disp(UnorderedMatrix);
            TruncatedOldSequance=TruncatedOldSequance(:,Temp_Seq);
        end
        Temp_Seq=1:UnarrangedMatrixSize;
        for Col=ZeroRows+1:UnarrangedMatrixSize
            AllZero=1;
            for Row=1:UnarrangedMatrixSize
                if (UnorderedMatrix(Row,Col)~=0 && Row~=Col)
                    AllZero=0;break;
                end
            end
            
            if (AllZero)
                ZeroColumns=ZeroColumns+1;
                Temp_Seq(Temp_Seq==Col)=[];
                Temp_Seq=[Temp_Seq,Col];
            end
        end
        if (ZeroColumns>0)
            Loop=1;
            UnorderedMatrix=UnorderedMatrix(:,Temp_Seq);
            UnorderedMatrix=UnorderedMatrix(Temp_Seq,:);
%             disp('Column');disp(UnorderedMatrix);
            TruncatedOldSequance=TruncatedOldSequance(:,Temp_Seq);
        end
        
        TempOrignalNewSequence=[TempOrignalNewSequence(1:TempFrom-1),TruncatedOldSequance,TempOrignalNewSequence(TempTo+1:end)];
        if (ZeroColumns>0 || ZeroRows>0)
            TempFrom=ZeroRows+1;TempTo=UnarrangedMatrixSize-ZeroColumns;
            UnorderedMatrix=UnorderedMatrix(TempFrom:TempTo,TempFrom:TempTo);
            TruncatedOldSequance=TruncatedOldSequance(:,TempFrom:TempTo);
        end
    end
end
OrderedMatrix=UnorderedMatrix;
OrignalNewSequence=TempOrignalNewSequence;
From=TempFrom;To=TempTo;
end

function [BlockAddress,OrignalNewSequence,IsValidCase]=BruteforceClustering(Matrix,OrignalOldSequence,MatrixFrom,MatrixTo)

TempTruncatedSequance=OrignalOldSequence(MatrixFrom:MatrixTo);
PermutationElements=1:size(Matrix,1);
PermutationList=perms(PermutationElements);
NumberOfPermutations=size(PermutationList,1);
%variable for storing best permutation
BestPermutationSequenceNumber=zeros(1,2);
InvalidPermutationSequence=zeros(1);
%checking for all possible permutation and picking the best one
ValidCase=1;
for PermutationNumber=1:NumberOfPermutations
    FromTo=zeros(1,2);TempMatrix=Matrix;
    TempMatrix=TempMatrix(:,PermutationList(PermutationNumber,:));
    TempMatrix=TempMatrix(PermutationList(PermutationNumber,:),:);
    MaxDist=zeros(1);Proceed=0;
    for Row=1:size(TempMatrix,1)
        for Column=size(TempMatrix,1):-1:Row
            if (Column==Row)
                MaxDist(Row)=0;
            elseif (TempMatrix(Row,Column)~=0)
                MaxDist(Row)=Column-Row;Proceed=1;
                break;
            end
        end
    end
    
    %for detection of all the blocks
    if (Proceed)
        Blocks=0;Row=1;Score=0;
        while (Row<=size(TempMatrix,1))
            if (MaxDist(Row)>0)
                From=Row;TempTo=From+1;To=Row+MaxDist(Row);TempMaxDist=From+MaxDist(From);
                while(To<=size(TempMatrix,1) && TempTo<=To)
                    if (MaxDist(TempTo)+TempTo>TempMaxDist)
                        To=TempTo+MaxDist(TempTo);
                        TempMaxDist=To;
                    end
                    TempTo=TempTo+1;
                end
                %calculating the percentage filled
                FilledCell=0;
                for k=From:To
                    for l=From:To
                        if (TempMatrix(k,l)~=0 && k~=l)
                            FilledCell=FilledCell+1;
                        end
                    end
                end
                PerFilled=(FilledCell*100/((To-From+1)^2-(To-From+1)));% percent filled
                %weighted scoring
                if (To-From+1==2)
                    Score=Score+0.15*PerFilled;
                elseif (To-From+1==3)
                    Score=Score+0.3*PerFilled;
                elseif (To-From+1==4)
                    Score=Score+0.8*PerFilled;
                elseif (To-From+1==5)
                    Score=Score+1.2*PerFilled;
                elseif (To-From+1>=5)
                    Score=Score+0.4*PerFilled;
                end
                Blocks=Blocks+1;
                FromTo(Blocks,1:2)=[From,To];
                Row=To+1;
                if (TempTo>size(TempMatrix,1))
                    break;
                end
            else
                Row=Row+1;
            end
        end
        if (BestPermutationSequenceNumber(1)<Score)
            BestPermutationSequenceNumber(1:2)=[Score,PermutationNumber];
            BestFromUpto=FromTo;
        end
    else
        InvalidPermutationSequence=PermutationNumber;
        ValidCase=0;break;
    end
    %for selecting the best permutation
    
end
if (ValidCase)
    BlockAddress=BestFromUpto;
    TempTruncatedSequance=TempTruncatedSequance(:,PermutationList(BestPermutationSequenceNumber(2),:));
    OrignalNewSequence=[OrignalOldSequence(1:MatrixFrom-1),TempTruncatedSequance,OrignalOldSequence(MatrixTo+1:end)];
    IsValidCase=1;
else
    BlockAddress=0;
    TempTruncatedSequance=TempTruncatedSequance(:,PermutationList(InvalidPermutationSequence(1),:));
    OrignalNewSequence=[OrignalOldSequence(1:MatrixFrom-1),TempTruncatedSequance,OrignalOldSequence(MatrixTo+1:end)];
    IsValidCase=0;
end
end