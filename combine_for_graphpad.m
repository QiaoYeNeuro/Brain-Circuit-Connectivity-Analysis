%%Combine CSI and PI numbers of individual mouse for data analysis
%%Qiao Ye, Xiangmin Xu lab, UCI, 04062023


%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

clear all;
%input counting file

[filename, path] = uigetfile({'*.xlsx','Excel Files(*.xlsx)'; '*.txt','Txt Files(*.txt)'}, 'Pick a file');
% load([path,'\',filename]);
% filename='Rabies SUB tracing-CSI&PI_03172023.csv';
% data order: 
% WT young M, WT mid M, AD young M, AD midM
% WT young F,WT mid F, AD young F,AD mid F,
% WT young=group1
% WT mid=group2
% AD young=group3
% AD mid=group4
[status,sheets]=xlsfinfo(filename);
% starter={};
 
%Region list%columns of brain regions picking for making AP plots
%bregma, 
Region_list={
'CA1_py'
'CA1_or'
'CA1_rad'
'CA1_lmol'
'CA2_py'
'CA2_or'
'CA3_py'
'CA3_or'
'SUB'
'postSUB'
'MS-DB'
'Thalamus'
'RSC'
'Vis Cortex'
'Aud Cortex'
'SS Cortex'
'TeA'
'Prh'
'Ect'
'LEC'
'MEC'
};


msnum=[8,10,10,12]; %mouse number in each group
 
% mouseind_raw=table2array(readtable(filename,'Sheet',2,'Range','C1:AP1')); 
% starter_raw=table2array(readtable(filename,'Sheet',2,'Range','C14:AP14'));
genderid_raw=table2array(readtable(filename,'Sheet',2,'Range','C2:AP2','ReadVariableNames',false));
CSI_raw=table2array(readtable(filename,'Sheet',2,'Range','C15:AP35','ReadVariableNames',false));
PI_raw=table2array(readtable(filename,'Sheet',2,'Range','C38:AP58','ReadVariableNames',false));
total_neuron_raw=table2array(readtable(filename,'Sheet',2,'Range','C4:AP4','ReadVariableNames',false));
connectivity_raw=table2array(readtable(filename,'Sheet',2,'Range','C5:AP5','ReadVariableNames',false));
starter_raw=table2array(readtable(filename,'Sheet',2,'Range','C14:AP14','ReadVariableNames',false));

genderid=nan(1,40);
for k=1:40   
genderid(k)=strcmp(genderid_raw(k),'M');
end

CSI_arranged=nan(21,max(msnum)*4);
CSI_arranged_MvF=nan(21,max(msnum)*8);
PI_arranged=nan(21,max(msnum)*4);
PI_arranged_MvF=nan(21,max(msnum)*8);
total_neuron_arranged=nan(1,max(msnum)*4);
total_neuron_arranged_MvF=nan(1,max(msnum)*8);
connectivity_arranged=nan(1,max(msnum)*4);
connectivity_arranged_MvF=nan(1,max(msnum)*8);
starter_arranged=nan(1,max(msnum)*4);
starter_arranged_MvF=nan(1,max(msnum)*8);

CSI_M=CSI_raw;
CSI_F=CSI_raw;

PI_M=PI_raw;
PI_F=PI_raw;

total_neuron_M=total_neuron_raw;
total_neuron_F=total_neuron_raw;

connectivity_M=connectivity_raw;
connectivity_F=connectivity_raw;

starter_M=starter_raw;
starter_F=starter_raw;


CSI_M(:,genderid(:)==0)=nan;
CSI_F(:,genderid(:)==1)=nan;
PI_M(:,genderid(:)==0)=nan;
PI_F(:,genderid(:)==1)=nan;
total_neuron_M(:,genderid(:)==0)=nan;
total_neuron_F(:,genderid(:)==1)=nan;
connectivity_M(:,genderid(:)==0)=nan;
connectivity_F(:,genderid(:)==1)=nan;
starter_M(:,genderid(:)==0)=nan;
starter_F(:,genderid(:)==1)=nan;
%%CSI----------------------
for i=1:length(msnum)
    if i==1
     CSI_arranged(:,((i-1)*max(msnum)+1):((i-1)*max(msnum)+msnum(i)))=CSI_raw(:,1:msnum(i));     
     CSI_arranged_MvF(:,((2*(i-1))*max(msnum)+1):((2*(i-1)*max(msnum)+msnum(i))))=CSI_M(:,1:msnum(i));%male
     CSI_arranged_MvF(:,((2*i-1)*max(msnum)+1):(((2*i-1)*max(msnum))+msnum(i)))=CSI_F(:,1:msnum(i));%female
    else
       CSI_arranged(:,((i-1)*max(msnum)+1):((i-1)*max(msnum)+msnum(i)))=CSI_raw(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
       CSI_arranged_MvF(:,((2*(i-1))*max(msnum)+1):((2*(i-1))*max(msnum)+msnum(i)))=CSI_M(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
       CSI_arranged_MvF(:,((2*i-1)*max(msnum)+1):(((2*i-1)*max(msnum))+msnum(i)))=CSI_F(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
    end
end

%%PI----------------------
for i=1:length(msnum)
    if i==1
     PI_arranged(:,((i-1)*max(msnum)+1):((i-1)*max(msnum)+msnum(i)))=PI_raw(:,1:msnum(i));     
     PI_arranged_MvF(:,((2*(i-1))*max(msnum)+1):((2*(i-1)*max(msnum)+msnum(i))))=PI_M(:,1:msnum(i));%male
     PI_arranged_MvF(:,((2*i-1)*max(msnum)+1):(((2*i-1)*max(msnum))+msnum(i)))=PI_F(:,1:msnum(i));%female
    else
       PI_arranged(:,((i-1)*max(msnum)+1):((i-1)*max(msnum)+msnum(i)))=PI_raw(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
       PI_arranged_MvF(:,((2*(i-1))*max(msnum)+1):((2*(i-1))*max(msnum)+msnum(i)))=PI_M(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
       PI_arranged_MvF(:,((2*i-1)*max(msnum)+1):(((2*i-1)*max(msnum))+msnum(i)))=PI_F(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
    end
end

%%Total neuron----------------------
for i=1:length(msnum)
    if i==1
     total_neuron_arranged(:,((i-1)*max(msnum)+1):((i-1)*max(msnum)+msnum(i)))=total_neuron_raw(:,1:msnum(i));     
     total_neuron_arranged_MvF(:,((2*(i-1))*max(msnum)+1):((2*(i-1)*max(msnum)+msnum(i))))=total_neuron_M(:,1:msnum(i));%male
     total_neuron_arranged_MvF(:,((2*i-1)*max(msnum)+1):(((2*i-1)*max(msnum))+msnum(i)))=total_neuron_F(:,1:msnum(i));%female
    else
       total_neuron_arranged(:,((i-1)*max(msnum)+1):((i-1)*max(msnum)+msnum(i)))=total_neuron_raw(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
       total_neuron_arranged_MvF(:,((2*(i-1))*max(msnum)+1):((2*(i-1))*max(msnum)+msnum(i)))=total_neuron_M(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
       total_neuron_arranged_MvF(:,((2*i-1)*max(msnum)+1):(((2*i-1)*max(msnum))+msnum(i)))=total_neuron_F(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
    end
end

%%Connecitivity----------------------
for i=1:length(msnum)
    if i==1
     connectivity_arranged(:,((i-1)*max(msnum)+1):((i-1)*max(msnum)+msnum(i)))=connectivity_raw(:,1:msnum(i));     
     connectivity_arranged_MvF(:,((2*(i-1))*max(msnum)+1):((2*(i-1)*max(msnum)+msnum(i))))=connectivity_M(:,1:msnum(i));%male
     connectivity_arranged_MvF(:,((2*i-1)*max(msnum)+1):(((2*i-1)*max(msnum))+msnum(i)))=connectivity_F(:,1:msnum(i));%female
    else
       connectivity_arranged(:,((i-1)*max(msnum)+1):((i-1)*max(msnum)+msnum(i)))=connectivity_raw(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
       connectivity_arranged_MvF(:,((2*(i-1))*max(msnum)+1):((2*(i-1))*max(msnum)+msnum(i)))=connectivity_M(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
       connectivity_arranged_MvF(:,((2*i-1)*max(msnum)+1):(((2*i-1)*max(msnum))+msnum(i)))=connectivity_F(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
    end
end

%%Starter----------------------
for i=1:length(msnum)
    if i==1
     starter_arranged(:,((i-1)*max(msnum)+1):((i-1)*max(msnum)+msnum(i)))=starter_raw(:,1:msnum(i));     
     starter_arranged_MvF(:,((2*(i-1))*max(msnum)+1):((2*(i-1)*max(msnum)+msnum(i))))=starter_M(:,1:msnum(i));%male
     starter_arranged_MvF(:,((2*i-1)*max(msnum)+1):(((2*i-1)*max(msnum))+msnum(i)))=starter_F(:,1:msnum(i));%female
    else
       starter_arranged(:,((i-1)*max(msnum)+1):((i-1)*max(msnum)+msnum(i)))=starter_raw(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
       starter_arranged_MvF(:,((2*(i-1))*max(msnum)+1):((2*(i-1))*max(msnum)+msnum(i)))=starter_M(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
       starter_arranged_MvF(:,((2*i-1)*max(msnum)+1):(((2*i-1)*max(msnum))+msnum(i)))=starter_F(:,sum(msnum(1:i-1))+1:sum(msnum(1:i-1))+msnum(i));
    end
end



    %Export files
    writematrix(CSI_arranged, [path,'\All_CSI','.csv']);
    writematrix(CSI_arranged_MvF, [path,'\MvF_CSI','.csv']);
    writematrix(PI_arranged, [path,'\All_PI','.csv']);
    writematrix(PI_arranged_MvF, [path,'\MvF_PI','.csv']);
    writecell([{'total neuron'},total_neuron_arranged], [path,'\All_sum_total_neuron','.csv']);
    writecell([{'total neuron'},total_neuron_arranged_MvF], [path,'\MvF_sum_total_neuron','.csv']);
    writecell([{'connectivity'},connectivity_arranged], [path,'\All_sum_connectivity','.csv']);
    writecell([{'connectivity'},connectivity_arranged_MvF], [path,'\MvF_sum_connectivity','.csv']);
    writecell([{'starter'},starter_arranged], [path,'\All_sum_starter','.csv']);
    writecell([{'starter'},starter_arranged_MvF], [path,'\MvF_sum_starter','.csv']);
    
   for r=1:size(CSI_raw(:,1))
          writecell([Region_list(r),CSI_arranged(r,:)], [path,num2str(r,'%02d '),'_All_CSI_',Region_list{r},'.csv']);
          writecell([Region_list(r),CSI_arranged_MvF(r,:)], [path,num2str(r,'%02d '),'_MvF_CSI_',Region_list{r},'.csv']);
          writecell([Region_list(r),PI_arranged(r,:)], [path,num2str(r,'%02d '),'_All_PI_',Region_list{r},'.csv']);
          writecell([Region_list(r),PI_arranged_MvF(r,:)], [path,num2str(r,'%02d '),'_MvF_PI_',Region_list{r},'.csv']);
   end


  