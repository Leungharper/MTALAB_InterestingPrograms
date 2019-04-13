% 2019.2.26————将图片写进Excel表格（一个单元格一个像素）
% 测试成功
clc; clear all; close all
filename = 'test3';

im = imread([filename '.jpg']);
% im(:,:,2)=95; %偏紫

[height width z] = size(im);
im = imresize(im,100/width); %压缩图片，防止循环过慢
im32 = double(im);
% 【将RGB分量（3个uint8的数）拼成24位的整数；但要改为double型】
color = im32(:,:,1) + im32(:,:,2)*2^8 + im32(:,:,3)*2^16;
% color = uint32(im(:,:,1)) + uint32(im(:,:,2))*256 + uint32(im(:,:,3))*2^16;
[height width] = size(color);

xaxis = tran(width); %横坐标
yaxis = cell(height,1); %纵坐标【构建元胞数组】
for i = 1:height
    yaxis(i) = {num2str(i)};
end

file = [filename '.xlsx'];
file = fullfile(pwd, file); % pwd：确定当前文件夹
h = actxserver('excel.application'); %Open an ActiveX connection to Excel

%Create a new work book (excel file)
wb=h.WorkBooks.Add();

for i = 1:height
    for j = 1:width
        ran = h.Activesheet.get('Range',[char(xaxis(j)),char(yaxis(i))]);
        ran.interior.Color = color(i,j); %【double型】
    end
end

% ran.interior.Color=hex2dec('000000'); %黑
% ran.interior.Color = 0; %黑

wb.SaveAs(file); %保存文件
wb.Close;
h.Quit;
h.delete;
    % This is a "Tab" test sentence!!
    % This is also a "Tab" test sentence!!
