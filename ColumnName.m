%for knowing the name of the column in excel sheet in general
function ColumnN=ColumnName(ExcelColumnNumber)
Text='ABCDEFGHIJKLMNOPQRSTUVWXYZ';
ExcelColumns=ExcelColumnNumber;
Quotient=fix(ExcelColumns/26);
NumberOfLetters=1;InitialExcludedColmns=26;
while (Quotient>1)
    Quotient=fix(Quotient/26);
    NumberOfLetters=NumberOfLetters+1;InitialExcludedColmns=InitialExcludedColmns+26^NumberOfLetters;
end
if (ExcelColumns>InitialExcludedColmns)
    NumberOfLetters=NumberOfLetters+1;
end
TempCount=1;
Letters=NumberOfLetters;
while (NumberOfLetters>1)
    ExcelColumns=ExcelColumns-26^TempCount;NumberOfLetters=NumberOfLetters-1;TempCount=TempCount+1;
end
%                 disp(ExcelColumns);
ColumnN=cell(1);ExcelColumns2=ExcelColumns;
for i=Letters:-1:1
    Quotient=fix(ExcelColumns2/26^(i-1));ExcelColumns=mod(ExcelColumns,26^(i-1));
    if (ExcelColumns2~=0)
        if (ExcelColumns~=0)
            ColumnN{1}=[ColumnN{1},Text(Quotient+1)];
        else
            ColumnN{1}=[ColumnN{1},Text(Quotient)];
        end
    else
        ColumnN{1}=[ColumnN{1},'Z'];
    end
    ExcelColumns2=ExcelColumns;
end
ColumnN=ColumnN{1};
end


