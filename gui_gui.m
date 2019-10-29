function varargout = gui_gui(varargin)
% GUI_GUI MATLAB code for gui_gui.fig
%      GUI_GUI, by itself, creates a new GUI_GUI or raises the existing
%      singleton*.
%
%      H = GUI_GUI returns the handle to a new GUI_GUI or the handle to
%      the existing singleton*.
%
%      GUI_GUI('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in GUI_GUI.M with the given input arguments.
%
%      GUI_GUI('Property','Value',...) creates a new GUI_GUI or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before gui_gui_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to gui_gui_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help gui_gui

% Last Modified by GUIDE v2.5 29-Oct-2019 09:44:01

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @gui_gui_OpeningFcn, ...
                   'gui_OutputFcn',  @gui_gui_OutputFcn, ...
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


% --- Executes just before gui_gui is made visible.
function gui_gui_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to gui_gui (see VARARGIN)

% Choose default command line output for gui_gui
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes gui_gui wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = gui_gui_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
javaFrame = get(gcf,'JavaFrame');
set(javaFrame,'Maximized',1);

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in pushbutton1_1.
function pushbutton1_1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton1_1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear file;
clear path;
clear file_path;
[file,path] = uigetfile('*.xlsx'); %浏览文件
file_suffix0 = file(end-5:end);
file_suffix = file_suffix0(strfind(file_suffix0,'.'):end) %判断文件类型
clear file_suffix0;
file_path = strcat(path,file)
if(file_suffix == '.xlsx' | file_suffix == '.xls') 
    xls_data=xlsread(file_path);    %读取Excel文件
    set(handles.uitable1,'Data',xls_data);
end

% --- Executes on button press in pushbutton1_2.
function pushbutton1_2_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton1_2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear table_data;clear data_cell;clear filter;
clear hangshu;clear lieshu;clear hang;clear lie;
clear Filename;clear Pathname;clear str;
clear CloumnName;clear dataExcel;

table_data = get(handles.uitable1,'Data');  %table_data含两个相同的元胞数组（2×1 cell 数组）
data_cell = cell2mat(table_data(1,1));    %转换元胞数组
filter = {'*.xlsx';'*.xls';'*.txt';'*.docx';'*.*'};
[Filename,Pathname] = uiputfile(filter,'另存为','data.xlsx');  %创建文件保存对话框
if (Filename==0 & Pathname==0)
	msgbox('您没有保存数据!','确认','error');
else
    str=[Pathname Filename];
    %获取表格的列名
    CloumnName=get(handles.uitable1,'ColumnName') ;                          
    CloumnName=CloumnName{2,1};
    
    dataExcel=cell(size(data_cell,1)+1,size(data_cell,2));
    dataExcel(1,:)=CloumnName;                                            %获取表格列名；
    dataExcel(2:end,:)=num2cell(data_cell);                                              %获取表格数据；
    xlswrite(str,dataExcel);                                              %将新单元数组写入指定的EXCEl文件中；
    
%     fid=fopen(str,'wt');    %重新打开建立的excel文件,可写
%     存数据
%     hangshu = size(data_cell,1);    %行数
%     lieshu = size(data_cell,2);     %列数
%     for hang=1:size(data_cell,1);    
%         for lie=1:size(data_cell,2) 
%             if(lie == size(data_cell,2))
%                  fprintf(fid,'%f\r',data_cell(hang,lie));
%             else
%                 fprintf(fid,'%f\t',data_cell(hang,lie));   
%             end
%         end
%     end
%     fclose(fid);        %关闭excel
    msgbox('保存数据完毕！','确认','warn');
end

% --- Executes on button press in pushbutton2_1.
function pushbutton2_1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton2_1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear file;
clear path;
clear file_path;
[file,path] = uigetfile('*.xlsx'); %浏览文件
file_suffix0 = file(end-5:end);
file_suffix = file_suffix0(strfind(file_suffix0,'.'):end);  %判断文件类型
clear file_suffix0;
file_path = strcat(path,file)
if(file_suffix == '.xlsx') 
    [~,~,xls_data]=xlsread(file_path);    %读取Excel文件
    set(handles.uitable2,'Data',xls_data);
end

% --- Executes on button press in pushbutton2_2.
function pushbutton2_2_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton2_2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear table_data;clear data_cell;clear filter;
clear hangshu;clear lieshu;clear hang;clear lie;
clear Filename;clear Pathname;clear str;
clear CloumnName;clear dataExcel;

table_data = get(handles.uitable2,'Data');  %table_data含两个相同的元胞数组（2×1 cell 数组）
data_cell = cell2mat(table_data(1,1));    %转换元胞数组
filter = {'*.xlsx';'*.xls';'*.txt';'*.docx';'*.*'};
[Filename,Pathname] = uiputfile(filter,'另存为','data.xlsx');  %创建文件保存对话框
if (Filename==0 & Pathname==0)
	msgbox('您没有保存数据!','确认','error');
else
    str=[Pathname Filename];
    %获取表格的列名
    CloumnName=get(handles.uitable2,'ColumnName') ;                          
    CloumnName=CloumnName{2,1};
    
    dataExcel=cell(size(data_cell,1)+1,size(data_cell,2));
    dataExcel(1,:)=CloumnName;                                            %获取表格列名；
    dataExcel(2:end,:)=num2cell(data_cell);                                              %获取表格数据；
    xlswrite(str,dataExcel);                                              %将新单元数组写入指定的EXCEl文件中；
    msgbox('保存数据完毕！','确认','warn');
end

% --- Executes on button press in pushbutton3_1.
function pushbutton3_1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton3_1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear file;
clear path;
clear file_path;
[file,path] = uigetfile('*.xlsx'); %浏览文件
file_suffix0 = file(end-5:end);
file_suffix = file_suffix0(strfind(file_suffix0,'.'):end);  %判断文件类型
clear file_suffix0;
file_path = strcat(path,file)
if(file_suffix == '.xlsx') 
    xls_data=xlsread(file_path);    %读取Excel文件
    set(handles.uitable3,'Data',xls_data);
end

% --- Executes on button press in pushbutton3_2.
function pushbutton3_2_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton3_2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear table_data;clear data_cell;clear filter;
clear hangshu;clear lieshu;clear hang;clear lie;
clear Filename;clear Pathname;clear str;
clear CloumnName;clear dataExcel;

table_data = get(handles.uitable3,'Data');  %table_data含两个相同的元胞数组（2×1 cell 数组）
data_cell = cell2mat(table_data(1,1));    %转换元胞数组
filter = {'*.xlsx';'*.xls';'*.txt';'*.docx';'*.*'};
[Filename,Pathname] = uiputfile(filter,'另存为','data.xlsx');  %创建文件保存对话框
if (Filename==0 & Pathname==0)
	msgbox('您没有保存数据!','确认','error');
else
    str=[Pathname Filename];
    %获取表格的列名
    CloumnName=get(handles.uitable3,'ColumnName') ;                          
    CloumnName=CloumnName{2,1};
    
    dataExcel=cell(size(data_cell,1)+1,size(data_cell,2));
    dataExcel(1,:)=CloumnName;                                            %获取表格列名；
    dataExcel(2:end,:)=num2cell(data_cell);                                              %获取表格数据；
    xlswrite(str,dataExcel);                                              %将新单元数组写入指定的EXCEl文件中；
    msgbox('保存数据完毕！','确认','warn');
end

% --- Executes on button press in pushbutton4_1.
function pushbutton4_1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton4_1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear file path file_path;
[file,path] = uigetfile('*.xlsx'); %浏览文件
file_suffix0 = file(end-5:end);
file_suffix = file_suffix0(strfind(file_suffix0,'.'):end);  %判断文件类型
clear file_suffix0;
file_path = strcat(path,file)
if(file_suffix == '.xlsx') 
    [num,txt,raw]=xlsread(file_path);    %读取Excel文件
    set(handles.uitable4,'Data',[raw]);
end

% --- Executes on button press in pushbutton4_2.
function pushbutton4_2_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton4_2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear table_data data_cell filter hangshu lieshu hang lie Filename Pathname str CloumnName dataExcel;

table_data = get(handles.uitable4,'Data');  %table_data含两个相同的元胞数组（2×1 cell 数组）
data_cell = cell2mat(table_data(1,1));    %转换元胞数组
filter = {'*.xlsx';'*.xls';'*.txt';'*.docx';'*.*'};
[Filename,Pathname] = uiputfile(filter,'另存为','data.xlsx');  %创建文件保存对话框
if (Filename==0 & Pathname==0)
	msgbox('您没有保存数据!','确认','error');
else
    str=[Pathname Filename];
    %获取表格的列名
    CloumnName=get(handles.uitable4,'ColumnName') ;                          
    CloumnName=CloumnName{2,1};
    
    dataExcel=cell(size(data_cell,1)+1,size(data_cell,2));
    dataExcel(1,:)=CloumnName;                                            %获取表格列名；
    dataExcel(2:end,:)=num2cell(data_cell);                                              %获取表格数据；
    xlswrite(str,dataExcel);                                              %将新单元数组写入指定的EXCEl文件中；
    msgbox('保存数据完毕！','确认','warn');
end

% --- Executes when entered data in editable cell(s) in uitable1.
function uitable1_CellEditCallback(hObject, eventdata, handles)
% hObject    handle to uitable1 (see GCBO)
% eventdata  structure with the following fields (see MATLAB.UI.CONTROL.TABLE)
%	Indices: row and column indices of the cell(s) edited
%	PreviousData: previous data for the cell(s) edited
%	EditData: string(s) entered by the user
%	NewData: EditData or its converted form set on the Data property. Empty if Data was not changed
%	Error: error string when failed to convert EditData to appropriate value for Data
% handles    structure with handles and user data (see GUIDATA)


% --- Executes when entered data in editable cell(s) in uitable5.
function uitable5_CellEditCallback(hObject, eventdata, handles)
% hObject    handle to uitable5 (see GCBO)
% eventdata  structure with the following fields (see MATLAB.UI.CONTROL.TABLE)
%	Indices: row and column indices of the cell(s) edited
%	PreviousData: previous data for the cell(s) edited
%	EditData: string(s) entered by the user
%	NewData: EditData or its converted form set on the Data property. Empty if Data was not changed
%	Error: error string when failed to convert EditData to appropriate value for Data
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in pushbutton12.
function pushbutton12_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton12 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear data_cell filter m  n hangshu lieshu hang  lie Filename Pathname str CloumnName dataExcel;

data_cell = get(handles.uitable5,'Data');    %转换元胞数组
[m,n] = size(data_cell);
if(m~=0 & n~=0)
    filter = {'*.docx';'*.*'};
    [Filename,Pathname] = uiputfile(filter,'另存为','data.docx');  %创建文件保存对话框
    if (Filename==0 & Pathname==0)
        msgbox('您没有保存数据!','确认','error');
    else
            str=[Pathname Filename];
            
            %生成Excel表格
%             CloumnName=get(handles.uitable5,'ColumnName');           %获取表格的列名             
%             dataExcel=cell(size(data_cell,1)+1,size(data_cell,2));
%             dataExcel(1,:)=CloumnName;                                            %存入表格列名；
%             dataExcel(2:end,:)=num2cell(data_cell);                                              %获取表格数据；
%             xlswrite(str,dataExcel);                                              %将新单元数组写入指定的EXCEl文件中；
%             msgbox('保存数据完毕！','确认','warn');

            %生成Word文档
            filespec_user = [str];  % 设定测试Word文件名和路径
            % 判断Word是否已经打开，若已打开，就在打开的Word中进行操作，否则就打开Word
            try
                % 若Word服务器已经打开，返回其句柄Word
                Word = actxGetRunningServer('Word.Application');
            catch
            % 否则，创建一个Microsoft Word服务器，返回句柄Word
                Word = actxserver('Word.Application');
            end;
            Word.Visible = 1; % 或set(Word, 'Visible', 1);

            % 若测试文件存在，打开该测试文件，否则，新建一个文件，并保存，文件名为测试.docx
            if exist(filespec_user,'file');
                Document = Word.Documents.Open(filespec_user);
            % Document = invoke(Word.Documents,'Open',filespec_user);
            else
                Document = Word.Documents.Add;
            % Document = invoke(Word.Documents, 'Add');
                Document.SaveAs2(filespec_user);
            end

            % 设定光标位置从头开始
            Content = Document.Content;
            Selection = Word.Selection;
            Paragraphformat = Selection.ParagraphFormat;

            % 设定页面大小
            Document.PageSetup.TopMargin = 60; % 单位像素
            Document.PageSetup.BottomMargin = 45;
            Document.PageSetup.LeftMargin = 45;
            Document.PageSetup.RightMargin = 45;

            % Content.InsertParagraphAfter;% 插入一段
            % Content.Collapse=0; % 0为不覆盖
            Content.Start = 0;
            title = '拉力扭矩预测';
            Content.Text = title;
            Content.Font.Size = 22 ;
            Content.Font.Bold = 4 ;
            Content.Paragraphs.Alignment = 'wdAlignParagraphCenter';% 设定段落格式
            Selection.Start = Content.end;% 定义开始的位置
            Selection.TypeParagraph;

            % 插入内容并定义字体字号
            smallTitle = '大庆油田有限责任公司试油试采分公司';
            Selection.Text = smallTitle;
            Selection.Font.Size = 12;
            Selection.Font.Bold = 0; 
            Selection.MoveDown;
            ParagraphFormat.Alignment = 'wdAlignParagraphCenter';
            Selection.TypeParagraph;    %插入新的空段落
            Selection.Font.Size = 10;

            %项目名称
            smallTitle = '项目名称：齐平1井钻桥塞';
            Selection.Text = smallTitle;
            Selection.Font.Size = 12;
            Selection.Font.Bold = 0; 
            Selection.MoveDown;
            Selection.ParagraphFormat.Alignment = 'wdAlignParagraphLeft';
            Selection.TypeParagraph;    %插入新的空段落
            
            %绘制表格
            Selection.ParagraphFormat.Alignment = 'wdAlignParagraphCenter'; %设置表格文字居中
            Selection.Font.Size = 10;   %设置表格中字体大小为10
            Tables = Document.Tables.Add(Selection.Range,size(data_cell,1),10);    % ”data_cell+1“ 行 10列
            DTI = Document.Tables.Item(1); % 或DTI = Tables;
            DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle';% 设置外边框的线型，Dash，DashDot,DashDotDot,DashSmallGap,DashLargeGap,Dot,Double,Triple等
            DTI.Borders.OutsideLineWidth = 'wdLineWidth150pt';% 设置线宽，有025，050，075，100，150，225，300，450，600pt等
            DTI.Borders.InsideLineStyle = 'wdLineStyleSingle';%设置内边框的线型
            DTI.Borders.InsideLineWidth = 'wdLineWidth075pt';
            DTI.Rows.Alignment = 'wdAlignRowCenter';%设置行对齐方式

            % DTI.Rows.Item(8).Borders.Item(1).LineStyle = 'wdLineStyleNone';% 设置第8行上边界线型，1，2，3，4分别对应单元格的上、左、下、右边界
            % DTI.Rows.Item(8).Borders.Item(3).LineStyle = 'wdLineStyleNone';% 设置第8行下边界线型
            % DTI.Rows.Item(11).Borders.Item(1).LineStyle = 'wdLineStyleNone';
            % DTI.Rows.Item(11).Borders.Item(3).LineStyle = 'wdLineStyleNone';

            column_width = [53.7736,85.1434,53.7736,35.0094,35.0094,76.6981,55.1887,52.9245,54.9057];% 设置列宽，单位为磅
            row_height = [28.5849,28.5849,28.5849,28.5849,25.4717,25.4717,32.8302,312.1698,17.8302,49.2453,14.1509,18.6792]; % 设置行高

            % 指定各单元格内容
            DTI.Cell(1,1).Range.Text = '序号';
            DTI.Cell(1,2).Range.Text = '测深(m)';
            DTI.Cell(1,3).Range.Text = '井斜角(°)';
            DTI.Cell(1,4).Range.Text = '方位角(°)';
            DTI.Cell(1,5).Range.Text = '拉力';
            DTI.Cell(1,6).Range.Text = '扭矩(kN·m)';
            DTI.Cell(1,7).Range.Text = '接触压(kN/m)';
            DTI.Cell(1,8).Range.Text = '稳定性';
            DTI.Cell(1,9).Range.Text = '安全系数';
            DTI.Cell(1,10).Range.Text = '伸长(m)';
%             DTI.Cell(1,10).Range.Font.Size = 10;

            %填写数据
            for cell_word = 1:size(data_cell,1)-1
                DTI.Cell(cell_word+1,1).Range.Text = num2str(cell_word);                            %序号 
                DTI.Cell(cell_word+1,2).Range.Text = num2str(data_cell{cell_word+1,1});      %测深    
                DTI.Cell(cell_word+1,3).Range.Text = num2str(data_cell{cell_word+1,2});     %井斜角
                DTI.Cell(cell_word+1,4).Range.Text = num2str(data_cell{cell_word+1,3});     %方位角
                DTI.Cell(cell_word+1,5).Range.Text = num2str(data_cell{cell_word+1,11});    %拉力   
                DTI.Cell(cell_word+1,6).Range.Text = num2str(data_cell{cell_word+1,12});    %扭矩  
                DTI.Cell(cell_word+1,7).Range.Text = num2str(data_cell{cell_word+1,13});    %接触压           
                DTI.Cell(cell_word+1,8).Range.Text = num2str(data_cell{cell_word+1,14});   %稳定性              
                DTI.Cell(cell_word+1,9).Range.Text = num2str(data_cell{cell_word+1,16});    %安全系数             
                DTI.Cell(cell_word+1,10).Range.Text = num2str(data_cell{cell_word+1,17});   %伸长
            end

            Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView'; % 设置视图方式为页面
            Document.Save; % 保存文档
%             Document.Close; % 关闭文档
%             Word.Quit; % 退出word服务器
    end
else
    msgbox('保存的数据不能为空！','确认','error');
end


% --- Executes on button press in pushbutton13.
function pushbutton13_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton13 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear table_data1;clear table_data2;clear table_data3;clear table_data4;
clear data_cell;clear data_cel2;clear data_cel3;clear data_cel4;
clear hang1;clear lie1;clear hang2;clear lie2;
clear hang3;clear lie3;clear hang4;clear lie4;
clear j;clear hang;

%测深&井斜角 计算
j = 1;
table_data1 = get(handles.uitable1,'Data');
data_cell = cell2mat(table_data1); 
for hang1=1:size(data_cell,1)
    table1_data1 = get(handles.uitable1(1,1),'Data');
    disp(table1_data1);
    data2 = table1_data1(1,1);
    data3 = table1_data1(2,1);
    data4 = table1_data1(1,2);
    data5 = table1_data1(2,2);
    fenshu = (data3 - data2)/10;
    lie_fen = (data5 - data4)/fenshu;
    disp(data4);
    disp(data5);
    disp(fenshu);
    disp(lie_fen);
    disp(data2);
    disp(data3);
    k = data4;
    for i = data2:10:data3
        table1_data(j,1) = i;
        table1_data(j,2) = k;
        %disp(table1_data);
        set(handles.uitable5(1,1),'Data',table1_data); 
        k = k + lie_fen;
        j = j+1;
    end
end

% %方位角
% 
% %地层温度
% 
% %井筒直径
% set(handles.uitable5(:,5),'Data','0.150'); 
% % for hang = 1:size(data_cell,1)
% %     set(handles.uitable5(hang,5),'Data','0.150'); 
% % end
% %管柱外径
% for hang = 1:size(data_cell,1)
%     set(handles.uitable5(hang,6),'Data','0.073'); 
% end
% %管柱内径
% for hang = 1:size(data_cell,1)
%     set(handles.uitable5(hang,7),'Data','0.057'); 
% end
% %环空流体密度
% for hang = 1:size(data_cell,1)
%     set(handles.uitable5(hang,8),'Data','1150.000'); 
% end
% %环空流体黏度
% 
% %管内流体
% 
% %轴向拉力
% 
% %扭矩
% for hang = 1:size(data_cell,1)
%     set(handles.uitable5(hang,12),'Data','0'); 
% end
% %接触压力
% 
% %稳定性状
% for hang = 1:size(data_cell,1)
%     set(handles.uitable5(hang,14),'Data','稳定'); 
% end
% %应力强度
% 
% %安全系数
% for hang = 1:size(data_cell,1)
%     set(handles.uitable5(hang,16),'Data','10.000'); 
% end
% %伸长

% --- Executes during object creation, after setting all properties.
function uipanel7_CreateFcn(hObject, eventdata, handles)
% hObject    handle to uipanel7 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called


% --- Executes during object creation, after setting all properties.
function axes1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to axes1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: place code in OpeningFcn to populate axes1


% --- Executes during object creation, after setting all properties.
function axes2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to axes2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: place code in OpeningFcn to populate axes2


% --- Executes during object creation, after setting all properties.
function axes3_CreateFcn(hObject, eventdata, handles)
% hObject    handle to axes3 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: place code in OpeningFcn to populate axes3


% --- Executes during object creation, after setting all properties.
function text11_CreateFcn(hObject, eventdata, handles)
% hObject    handle to text11 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called


% --- Executes on button press in pushbutton14.
function pushbutton14_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton14 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear file;
clear path;
clear file_path;
[file,path] = uigetfile('*.xlsx'); %浏览文件
file_suffix0 = file(end-5:end);
file_suffix = file_suffix0(strfind(file_suffix0,'.'):end);  %判断文件类型
clear file_suffix0;
file_path = strcat(path,file)
if(file_suffix == '.xlsx') 
    [num,txt,raw]=xlsread(file_path);    %读取Excel文件
    set(handles.uitable5,'Data',[raw]);
end


% --- Executes when selected cell(s) is changed in uitable1.
function uitable1_CellSelectionCallback(hObject, eventdata, handles)
% hObject    handle to uitable1 (see GCBO)
% eventdata  structure with the following fields (see MATLAB.UI.CONTROL.TABLE)
%	Indices: row and column indices of the cell(s) currently selecteds
% handles    structure with handles and user data (see GUIDATA)
%行列索引
global row;
global col;
index=eventdata.Indices
row=index(1);
col=index(2);


% --- Executes on button press in pushbutton15.
function pushbutton15_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton15 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in pushbutton16.
function pushbutton16_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton16 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear newData;
global row col;
if(row>0)
    newData = get(handles.uitable1,'Data');  %获取表格数据矩阵
    newData(row,:) = [];   %删除选中的某行数据
    set(handles.uitable1,'Data',newData);  %显示到表格中
else
    msgbox('请先选择要删除的行！','确认','error');
end

% --- Executes on button press in pushbutton17.
function pushbutton17_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton17 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in pushbutton18.
function pushbutton18_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton18 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear newData;
global row col;
if(row>0)
    newData = get(handles.uitable1,'Data');  %获取表格数据矩阵
    newData(row,:) = [];   %删除选中的某行数据
    set(handles.uitable1,'Data',newData);  %显示到表格中
else
    msgbox('请先选择要删除的行！','确认','error');
end

% --- Executes on button press in pushbutton19.
function pushbutton19_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton19 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear newData;
global row col;
if(row>0)
    newData = get(handles.uitable1,'Data');  %获取表格数据矩阵
    newData(row,:) = [];   %删除选中的某行数据
    set(handles.uitable1,'Data',newData);  %显示到表格中
else
    msgbox('请先选择要删除的行！','确认','error');
end

% --- Executes on button press in pushbutton20.
function pushbutton20_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton20 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in pushbutton21.
function pushbutton21_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton21 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in pushbutton22.
function pushbutton22_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton22 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
clear newData;
global row col;
if(row>0)
    newData = get(handles.uitable1,'Data');  %获取表格数据矩阵
    newData(row,:) = [];   %删除选中的某行数据
    set(handles.uitable1,'Data',newData);  %显示到表格中
else
    msgbox('请先选择要删除的行！','确认','error');
end


% --- Executes on button press in pushbutton23.
function pushbutton23_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton23 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
table_data3 = get(handles.uitable5(1,1),'Data');
data3_1 = table_data3(:,1);
data3_2 = table_data3(:,2);
x = cell2mat(data3_1);
y = cell2mat(data3_2);
disp(x);
disp(y);
axes(handles.axes1); 
plot(x,y);

% --- Executes on button press in pushbutton24.
function pushbutton24_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton24 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in pushbutton25.
function pushbutton25_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton25 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
