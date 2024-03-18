function varargout = phyloscript(varargin) % PHYLOSCRIPT MATLAB code for phyloscript.fig
 gui_Singleton=1; % Begin initialization code - DO NOT EDIT
 gui_State=struct('gui_Name', mfilename, 'gui_Singleton', gui_Singleton, 'gui_OpeningFcn', @phyloscript_OpeningFcn, ...
                   'gui_OutputFcn', @phyloscript_OutputFcn, 'gui_LayoutFcn', [], 'gui_Callback', []);
 if nargin && ischar(varargin{1}) gui_State.gui_Callback=str2func(varargin{1}); end
 if nargout [varargout{1:nargout}]=gui_mainfcn(gui_State, varargin{:}); else gui_mainfcn(gui_State, varargin{:}); end
end % End initialization code - DO NOT EDIT

function phyloscript_OpeningFcn(hObject, eventdata, handles, varargin) % Executes just before phyloscript is made visible
  global dissimilarityType dissimilarityNameType linkageType linkageNameType objectType plotType;
  handles.fileDataMatrix='datamatrix.xlsx'; % Input data in Excel
  handles.scriptClassRange='A2:A5'; handles.scriptNameRange='B2:B5'; handles.scriptFeatureRange='C2:DQ5';
  handles.inscriptionClassRange='A6:A62'; handles.inscriptionNameRange='B6:B62'; handles.inscriptionFeatureRange='C6:DQ62';
  handles.featureNameRange='C1:DQ1';
  objectType='inscriptions';dissimilarityType='sorensen-dice';linkageType='weighted';linkageNameType='WPGMA';
  plotType='square';handles.objectNamesFontSize=8;handles.markerSize=20;dissimilarityNameType='Sørensen-Dice dissimilarity';
  handles.output=hObject; % Choose default command line output for phyloscript
  resetDataSet(hObject,eventdata,handles);guidata(hObject,handles); % Call reset dataset & update handles structure
end

function varargout=phyloscript_OutputFcn(~,~,handles) % Outputs from this function are returned to the command line
    varargout{1}=handles.output;
end

function resetDataSet(~,~,handles) % Reset dataset
 global featureF featureNameVector objectClassVector objectN objectNameVector objectType scriptNames XMatrix;
 [~,scriptNames]=(xlsread(handles.fileDataMatrix,handles.scriptClassRange)); % Load script (class) names into a vector
 if strcmp(objectType,'inscriptions') % The objects are inscriptions
  XMatrix=(xlsread(handles.fileDataMatrix,handles.inscriptionFeatureRange)); % Load inscription data to XMatrix, row:inscription, column:feature
  [~,objectClassVector]=(xlsread(handles.fileDataMatrix,handles.inscriptionClassRange)); % Load inscription classes into objectClassVector
  [~,objectNameVector]=(xlsread(handles.fileDataMatrix,handles.inscriptionNameRange)); % Load inscription names into objectNameVector
  [~,featureNameVector]=(xlsread(handles.fileDataMatrix,handles.featureNameRange)); % Load feature names into featureNameVector
 end
 if strcmp(objectType,'scripts') % The objects are scripts
  XMatrix=(xlsread(handles.fileDataMatrix,handles.scriptFeatureRange)); % Load script data to XMatrix, row:script, column:feature
  [~,objectClassVector]=(xlsread(handles.fileDataMatrix,handles.scriptClassRange)); % Load script classes into objectClassVector
  [~,objectNameVector]=(xlsread(handles.fileDataMatrix,handles.scriptNameRange)); % Load script names into objectNameVector
  [~,featureNameVector]=(xlsread(handles.fileDataMatrix,handles.featureNameRange)); % Load feature names into featureNameVector
 end
 XMatrix(isnan(XMatrix))=0; % If there are not a number elements, they will be reset
 objectN=size(objectNameVector'); featureF=size(featureNameVector); % Calculate # of objects and features after preprocessing
end

function HCA_Callback(~,~,handles)
 global dissimilarityNameType featureF linkageNameType linkageType objectN objectNameVector;
 D=modifiedpdist; Z = linkage(D,linkageType); % Calculating dissimilarities then linkages (a tree)
 [cpcc,~]=cophenet(Z,D); % Calculating the cophenetic correlation coefficient and the cophenetic dissimilarities
 figure() % Creating a new window for the diagram
 leafOrder=optimalleaforder(Z,D); % Leaf ordering
 [H,~]=dendrogram(Z,0,'reorder',leafOrder,'labels', objectNameVector','Orientation','left');
 set(gca,'fontsize',handles.objectNamesFontSize);set(H,'LineWidth',2.2,'color',[0 0 0]);
 xlabel([linkageNameType,' linkage, ',dissimilarityNameType,', ',num2str(objectN(2)),' objects, ',num2str(featureF(2)),' features, CPCC=', ...
     num2str(cpcc,'%.4f')]);
end

function NJ_Callback(~,~,handles)
 global dissimilarityNameType featureF objectN objectNameVector plotType;
 D=modifiedpdist; % Call dissimilarity calculation
 phytree=seqneighjoin(D,'equivar',objectNameVector);
 h=plot(phytree,'type',plotType); %view(phytree);
 set(gca,'FontSize',handles.objectNamesFontSize);set(h.terminalNodeLabels,'FontSize',handles.objectNamesFontSize);
 xlabel(['NJ tree, ',dissimilarityNameType,', ',num2str(objectN(2)),' objects and ',num2str(featureF(2)),' features']);
end

function PCA2_Callback(~,~,handles)
 global objectClassVector objectN scriptNames XMatrix; [~,score,latent,~,explained,~]=pca(XMatrix); figure()
 for i=1:objectN(2)
  if strcmp(objectClassVector(i),'TR'),pClassTR=plot(score(i,1),score(i,2),'.','markerSize',handles.markerSize,'color',[0 0 1]);hold on,end %blue
  if strcmp(objectClassVector(i),'SHR'),pClassSHR=plot(score(i,1),score(i,2),'.','markerSize',handles.markerSize,'color',[0 1 0]);end %green
  if strcmp(objectClassVector(i),'CBR'),pClassCBR=plot(score(i,1),score(i,2),'.','markerSize',handles.markerSize,'color',[1 0.6 0]);end %orange
  if strcmp(objectClassVector(i),'SR'),pClassSR=plot(score(i,1),score(i,2),'.','markerSize',handles.markerSize,'color',[1 0 1]);end %violet/magenta 
 end
 axis equal; xlabel('1st Principal Component'); ylabel('2nd Principal Component')
 legend([pClassTR pClassSHR pClassCBR pClassSR],scriptNames,'Location','NE'); grid on; hold off;
 figure(); LastName = {'1st Principal Component';'2nd Principal Component'};
 Eigenvalue=[latent(1);latent(2)];Variance=[explained(1);explained(2)];T= table(Eigenvalue,Variance,'RowNames',LastName);
 uitable('Data',T{:,:},'ColumnName',T.Properties.VariableNames,...
    'RowName',T.Properties.RowNames,'Units', 'Normalized', 'Position',[0, 0, 1, 1]);end

function PCA3_Callback(~,~,handles) % Executes on button press in PCA3
 global objectClassVector objectN scriptNames XMatrix;
 [~,score,latent,~,explained,~] = pca(XMatrix); figure()
 for i=1:objectN(2)
  if strcmp(objectClassVector(i),'TR'),pClassTR=plot3(score(i,1),score(i,2),score(i,3),'.','Color',[0 0 1],'MarkerSize',handles.markerSize);hold on,end %blue
  if strcmp(objectClassVector(i),'SHR'),pClassSHR=plot3(score(i,1),score(i,2),score(i,3),'.','Color',[0 1 0],'MarkerSize',handles.markerSize);end %green
  if strcmp(objectClassVector(i),'CBR'),pClassCBR=plot3(score(i,1),score(i,2),score(i,3),'.','Color',[1 0.6 0],'MarkerSize',handles.markerSize);end %orange
  if strcmp(objectClassVector(i),'SR'),pClassSR=plot3(score(i,1),score(i,2),score(i,3),'.','Color',[1 0 1],'MarkerSize',handles.markerSize);end %violet/magenta
 end
 axis equal; xlabel('1st Principal Component'); ylabel('2nd Principal Component'); zlabel('3rd Principal Component');
 legend([pClassTR pClassSHR pClassCBR pClassSR],scriptNames,'Location','NE'); grid on; hold off;
 figure(); LastName = {'1st Principal Component';'2nd Principal Component';'3rd Principal Component'};
 Eigenvalue=[latent(1);latent(2);latent(3)];Variance=[explained(1);explained(2);explained(3)];T= table(Eigenvalue,Variance,'RowNames',LastName);
 uitable('Data',T{:,:},'ColumnName',T.Properties.VariableNames,...
    'RowName',T.Properties.RowNames,'Units', 'Normalized', 'Position',[0, 0, 1, 1]);
end

function objectSelector_Callback(hObject,eventdata,handles) % Executes on selection change in objectSelector
  global objectType; objectType='scripts'; objectVal=get(handles.objectSelector,'Value');
  switch objectVal, case 2, objectType='scripts'; case 3, objectType='inscriptions'; end
  resetDataSet(hObject, eventdata, handles); guidata(hObject, handles); % Reload Excel datasheet & update handles structure
end

function dissimilaritySelector_Callback(hObject, ~, handles) % Executes on selection change in dissimilaritySelector
 global dissimilarityNameType dissimilarityType;dissimilarityType='sorensen-dice';dissimilarityVal=get(handles.dissimilaritySelector,'Value');
 switch dissimilarityVal
     case 2, dissimilarityType='sorensen-dice';dissimilarityNameType='Sørensen-Dice dissimilarity';
     case 3, dissimilarityType='jaccard';dissimilarityNameType='Jaccard distance';
     case 4, dissimilarityType='cosine';dissimilarityNameType='Cosine distance';     
     case 5, dissimilarityType='euclidean';dissimilarityNameType='Euclidean distance';
 end
 guidata(hObject,handles); % Update handles structure
end

function linkageSelector_Callback(hObject,~,handles) % Executes on linkage selection change
 global linkageNameType linkageType;linkageVal=get(handles.linkageSelector,'Value');
 switch linkageVal
   case 2, linkageType='single';linkageNameType='Single (nearest neighbour)';
   case 3, linkageType='complete';linkageNameType='Complete (farthest neighbour)';
   case 4, linkageType='average';linkageNameType='UPGMA';
   case 5, linkageType='weighted';linkageNameType='WPGMA kutykurutty';
   case 6, linkageType='centroid';linkageNameType='UPGMC (centroid)';
   case 7, linkageType='median';linkageNameType='WPGMC (median)';
   case 8, linkageType='ward';linkageNameType='Ward (minimum variance)';
 end
 guidata(hObject,handles); % Update handles structure
end

function plotType_Callback(hObject,~,handles) % Executes on plot type for NJ selection change
 global plotType;plotTypeVal=get(handles.plotType,'Value');
 switch plotTypeVal
   case 2, plotType='square';
   case 3, plotType='angular';
   case 4, plotType='radial';
   case 5, plotType='equalangle';
   case 6, plotType='equaldaylight';
 end
 guidata(hObject,handles); % Update handles structure
end

function D=modifiedpdist % Calculating Sørensen-Dice or other kind of dissimilarity, the latter uses the built-in Matlab function pdist
 global dissimilarityType XMatrix;
 if strcmp(dissimilarityType,'sorensen-dice'), D=dice(XMatrix); else, D=pdist(XMatrix, dissimilarityType); end
end

function D=dice(XMatrix) % Sørensen-Dice dissimilarity calculation, D: output dissimilarity matrix
  [rowN,columnN]=size(XMatrix); % rowN: number of rows (=objectN), columnN: number of columns (=featureF)
  DMatrix=zeros(rowN,rowN); % Preallocating array DMatrix to avoid changing its size in each iteration to increase speed
  for i=1:rowN % Selecting one object
   for j=1:rowN % Selecting another object
      a=0; b=0; c=0; %a: both 1, b & c: one is 0, the other is 1 and opposite case, d: both are 0 (unused)
      for k=1:columnN % Comparing the corresponding features of the selected objects
        if XMatrix(i,k)==1 && XMatrix(j,k)==1, a=a+1; end % Calculation of number of features being equally 1
        if XMatrix(i,k)==1 && XMatrix(j,k)==0, b=b+1; end % Calculation of number of features with alternate values
        if XMatrix(i,k)==0 && XMatrix(j,k)==1, c=c+1; end % Calculation of number of features with alternate values in opposite case
      end
      DMatrix(i,j)=1-(2*a)/(2*a+b+c);
   end
  end
  D=squareform(DMatrix); % Converting the square dissimilarity matrix DMatrix into D, a vector containing the DMatrix elements below the diagonal
end