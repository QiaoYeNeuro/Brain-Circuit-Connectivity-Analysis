%%Combine CSI numbers of individual mouse and align AP number
%%Qiao Ye, 02232023


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

clear all;
%input counting file

[filename, path] = uigetfile({'*.xlsx','Excel Files(*.xlsx)'; '*.txt','Txt Files(*.txt)'}, 'Pick a file');
% load([path,'\',filename]);
% filename='Rabies SUB tracing-CSI&PI_02232023.xlsx';
% data order: 
% WT young M, WT mid M, AD young M, AD midM
% WT young F,WT mid F, AD young F,AD mid F,
% WT young=group1
% WT mid=group2
% AD young=group3
% AD mid=group4
[status,sheets]=xlsfinfo(filename);
starter={};
sheetnum={};

columnlist={%columns of brain regions picking for making AP plots
%bregma, 
% CA1_py
% CA1_or
% CA1_rad
% CA1_lmol
% CA2_py
% CA2_or
% CA3_py
% CA3_or
% SUB
% postSUB
% MS-DB
% Thalamus
% RSC
% Vis Cortex
% Aud Cortex
% SS Cortex
% TeA
% Prh 
% Ect
% LEC
% MEC
'C4:C80'
'Y4:Y80'
'Z4:Z80'
'AA4:AA80'
'AB4:AB80'
'AC4:AC80'
'AD4:AD80'
'AF4:AF80'
'AG4:AG80'
'H4:H80'
'I4:I80'
'K4:K80'
'O4:O80'
'L4:L80'
'P4:P80'
'Q4:Q80'
'R4:R80'
'S4:S80'
'T4:T80'
'U4:U80'
'W4:W80'
'X4:X80'
};


msnum=[8,10,10,12]; %mouse number in each group
 
mouseind_raw=table2array(readtable(filename,'Sheet',2,'Range','C1:AP1')); 
starter_raw=table2array(readtable(filename,'Sheet',2,'Range','C14:AP14'));
genderid_raw=table2array(readtable(filename,'Sheet',2,'Range','C2:AP2','ReadVariableNames',false));

tl=cell(length(msnum),max(msnum));
tl(:)={NaN(100,length(columnlist))};
tl_group=cell(length(msnum),1);
tl_group(:)={NaN(100,length(columnlist)*max(msnum))};
tl_group_gender=cell(length(msnum),2);
tl_group_gender(:)={NaN(100,length(columnlist)*max(msnum))};
k=1;
for i=1:length(msnum) %number of mouse groups
%         group_tl{i}=nan(100,15*msnum(i)); %100 is maximal row number, %15 is maximal column number for each group

    for j=1:msnum(i)
        starter{i,j}=starter_raw(sum(msnum(1:i-1))+j);%subiculum starter neuron of each mouse
        mouseind{i,j}=mouseind_raw(sum(msnum(1:i-1))+j);%mouse index 
        genderid{i,j}=strcmp(genderid_raw(sum(msnum(1:i-1))+j),'F');
        sheetnum{i,j}=find(~cellfun(@isempty,strfind(sheets(:),num2str(mouseind{i,j},'%02d '))));%sheet number of the j-th mouse brain in 'sheets'
      
        for c=1:length(columnlist) 
        %extract data from excel sheet of that column
            columndata=readtable(filename,'Sheet',sheetnum{i,j},'Range',columnlist{c},'Format','auto');  %extract data from excel sheet of that column
%             columndata{~cellfun(@isempty,columndata)}=nan;
             
            tl{i,j}(1:length(table2array(columndata)),c)=table2array(columndata);
              
                if c==1
                    tl_group{i}(1:length(table2array(columndata)),(max(msnum)*(c-1)+j))=table2array(columndata);
                else
                    tl_group{i}(1:length(table2array(columndata)),(max(msnum)*(c-1)+j))=table2array(columndata)/starter{i,j};
                end
                %Separate genders
                 if c==1
                     tl_group_gender{i,genderid{i,j}+1}(1:length(table2array(columndata)),(max(msnum)*(c-1)+j))=table2array(columndata);
                 else
                    tl_group_gender{i,genderid{i,j}+1}(1:length(table2array(columndata)),(max(msnum)*(c-1)+j))=table2array(columndata)/starter{i,j};
                 end
        end
        fprintf('%d brains completed',k);
        k=k+1;        
    end

    %Replace zeros with empty nan
    tl_group{i,1}((iszero(tl_group{i,1}(:))))=nan;
    tl_group_gender{i,1}((iszero(tl_group_gender{i,1}(:))))=nan;
    tl_group_gender{i,2}((iszero(tl_group_gender{i,2}(:))))=nan;


    %Match AP positions
    tl_group_adjusted{i,1}=tl_group{i,1};
    tl_group_gender_adjusted{i,1}=tl_group_gender{i,1};
    tl_group_gender_adjusted{i,2}=tl_group_gender{i,2};
    
    %AP number of each group
    AP{i,1}=tl_group{i,1}(1,1:max(msnum));
    AP_gender{i,1}=tl_group_gender{i,1}(1,1:max(msnum));
    AP_gender{i,2}=tl_group_gender{i,2}(1,1:max(msnum));
    %Align AP number of entire group

    for l=find(~isnan(AP{i,1}))%loop in between non empty columns
        rowshift{i,l}=round((max(tl_group{i,1}(1,1:max(msnum)))-tl_group{i,1}(1,l))/0.09);%how many rows to shift down , 0.09 is brain slice gap in mm.        
        for m=1:length(columnlist) 
            tl_group_adjusted{i,1}(:,l+max(msnum)*(m-1))=circshift(tl_group_adjusted{i,1}(:,l+max(msnum)*(m-1)),rowshift{i,l});
        end
    end
    
    %Align AP number of entire male group
    for l=find(~isnan(AP_gender{i,1}))%loop in between non empty columns
        rowshift_gender{i,1}(l)=round((max(tl_group_gender{i,1}(1,1:max(msnum)))-tl_group_gender{i,1}(1,l))/0.09);%how many rows to shift down , 0.09 is brain slice gap in mm.        
        for m=1:length(columnlist) %loop through all brain regions
            tl_group_gender_adjusted{i,1}(:,l+max(msnum)*(m-1))=circshift(tl_group_gender_adjusted{i,1}(:,l+max(msnum)*(m-1)),rowshift_gender{i,1}(l));
        end
    end
 
    %Align AP number of entire female group
    for l=find(~isnan(AP_gender{i,2}))%loop in between non empty columns
        rowshift_gender{i,2}(l)=round((max(tl_group_gender{i,2}(1,1:max(msnum)))-tl_group_gender{i,2}(1,l))/0.09);%how many rows to shift down , 0.09 is brain slice gap in mm.        
        for m=1:length(columnlist) %loop through all brain regions
            tl_group_gender_adjusted{i,2}(:,l+max(msnum)*(m-1))=circshift(tl_group_gender_adjusted{i,2}(:,l+max(msnum)*(m-1)),rowshift_gender{i,2}(l));
        end
    end

    
    %Export files
    writematrix(tl_group{i,1}, ['C:\Users\exx\Desktop\grouped_',num2str(i),'.xlsx']);
    writematrix(tl_group_gender{i,1}, [path,'\grouped_male_ ',num2str(i),'.xlsx']);
    writematrix(tl_group_gender{i,2}, [path,'\grouped_female_ ',num2str(i),'.xlsx']);
    
    writematrix(tl_group_adjusted{i,1}, [path,'\adjsuted_grouped_',num2str(i),'.xlsx']);
    writematrix(tl_group_gender_adjusted{i,1}, [path,'\adjusted_grouped_male_ ',num2str(i),'.xlsx']);
    writematrix(tl_group_gender_adjusted{i,2}, [path,'\adjusted_grouped_female_ ',num2str(i),'.xlsx']);

end
      
 
 
        
        
