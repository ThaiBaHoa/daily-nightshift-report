import React, { useState, useEffect, useRef } from 'react';
import { 
  Button, 
  TextField, 
  Container, 
  Box, 
  Typography,
  Paper,
  Grid,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  FormHelperText,
  Stack,
  Snackbar,
  Alert,
  CircularProgress,
  IconButton,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  ImageList,
  ImageListItem
} from '@mui/material';
import { AdapterDateFns } from '@mui/x-date-pickers/AdapterDateFns';
import { LocalizationProvider, DatePicker } from '@mui/x-date-pickers';
import * as XLSX from 'xlsx';
import { format, parse } from 'date-fns';
import AddPhotoAlternateIcon from '@mui/icons-material/AddPhotoAlternate';
import DeleteIcon from '@mui/icons-material/Delete';
import VisibilityIcon from '@mui/icons-material/Visibility';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { FIELD_INSTRUCTIONS, getFieldInstruction } from './fieldInstructions';

interface TemplateField {
  value: string | number | null;
  isEditable: boolean;
}

interface TemplateRow {
  [key: string]: TemplateField;
}

interface ImageAttachment {
  dataUrl: string;
  name: string;
}

interface DataRow {
  [key: string]: string | number | null | ImageAttachment[];
}

interface SelectChangeEvent {
  target: {
    value: string;
  };
}

const INSPECTORS = [
  "TBHOA",
  "DTPHU",
  "CTHUY",
  "NQHUY",
  "NGHTHO",
  "LPTPHONG",
  "LHQUAN",
  "TVHOANG",
  "TTTHIEN"
];

const STATIONS = [
  "SGN",
  "HAN",
  "DAD"
];

const STATUS_OPTIONS = [
  "Checked",
  "Not Check",
  "Finding"
];

function App() {
  const [data, setData] = useState<DataRow[]>([]);
  const [template, setTemplate] = useState<TemplateRow>({});
  const [headers, setHeaders] = useState<string[]>([]);
  const [selectedInspector, setSelectedInspector] = useState<string>('');
  const [selectedStation, setSelectedStation] = useState<string>('');
  const [selectedDate, setSelectedDate] = useState<Date>(new Date());
  const [selectedSTT, setSelectedSTT] = useState<number>(1);
  const [excelFormat, setExcelFormat] = useState<any>(null);
  const [loading, setLoading] = useState<boolean>(false);
  const [snackbar, setSnackbar] = useState<{
    open: boolean;
    message: string;
    severity: 'success' | 'info' | 'warning' | 'error';
  }>({
    open: false,
    message: '',
    severity: 'info'
  });
  
  // Image attachment states
  const [imagePreviewOpen, setImagePreviewOpen] = useState<boolean>(false);
  const [selectedImage, setSelectedImage] = useState<string>('');
  const fileInputRef = useRef<HTMLInputElement>(null);
  
  const MAX_IMAGE_WIDTH = 800;
  const MAX_IMAGE_HEIGHT = 600;

  const showSnackbar = (message: string, severity: 'success' | 'info' | 'warning' | 'error' = 'info') => {
    setSnackbar({
      open: true,
      message,
      severity
    });
  };

  const handleCloseSnackbar = () => {
    setSnackbar({
      ...snackbar,
      open: false
    });
  };

  const formatDate = (date: Date): string => {
    try {
      if (!(date instanceof Date) || isNaN(date.getTime())) {
        console.warn('Invalid date provided, using current date as fallback');
        return format(new Date(), 'dd-MM-yyyy');
      }
      return format(date, 'dd-MM-yyyy');
    } catch (error) {
      console.error('Error formatting date:', error);
      return format(new Date(), 'dd-MM-yyyy');
    }
  };

  const parseDate = (dateString: string): Date => {
    try {
      // Hỗ trợ nhiều định dạng ngày tháng phổ biến
      const formats = ['dd-MM-yyyy', 'yyyy-MM-dd', 'MM/dd/yyyy'];
      
      for (const formatStr of formats) {
        try {
          const parsedDate = parse(dateString, formatStr, new Date());
          if (!isNaN(parsedDate.getTime())) {
            return parsedDate;
          }
        } catch (e) {
          // Thử định dạng tiếp theo
        }
      }
      
      // Nếu không thể parse, trả về ngày hiện tại
      console.warn(`Không thể parse chuỗi ngày: ${dateString}, sử dụng ngày hiện tại`);
      return new Date();
    } catch (error) {
      console.error('Error parsing date:', error);
      return new Date();
    }
  };

  useEffect(() => {
    loadTemplateFile();
    loadTempFile();
  }, []);

  const deleteTempFile = () => {
    try {
      localStorage.removeItem('tempData');
      showSnackbar('Đã xóa dữ liệu tạm thành công!', 'success');
    } catch (error) {
      console.error('Error deleting temp data:', error);
      showSnackbar('Không thể xóa dữ liệu tạm!', 'error');
    }
  };

  const saveTempFile = () => {
    if (!selectedInspector) {
      showSnackbar('Vui lòng nhập tên INSPECTOR trước khi lưu!', 'warning');
      return;
    }

    if (!selectedStation) {
      showSnackbar('Vui lòng nhập tên STATION trước khi lưu!', 'warning');
      return;
    }

    try {
      // Lưu dữ liệu vào localStorage thay vì file tạm
      const tempData = {
        data,
        inspector: selectedInspector,
        station: selectedStation,
        date: selectedDate.toISOString()
      };
      
      // Kiểm tra kích thước dữ liệu trước khi lưu
      const tempDataString = JSON.stringify(tempData);
      const dataSize = new Blob([tempDataString]).size;
      
      // Nếu dữ liệu quá lớn (> 5MB), hiển thị cảnh báo
      if (dataSize > 5 * 1024 * 1024) {
        console.warn('Dữ liệu tạm quá lớn, có thể gây vấn đề với localStorage');
        showSnackbar('Dữ liệu tạm quá lớn, một số hình ảnh có thể không được lưu!', 'warning');
      }
      
      localStorage.setItem('tempData', tempDataString);
      console.log('Đã lưu dữ liệu tạm thời thành công');
      showSnackbar('Đã lưu dữ liệu tạm thành công!', 'success');
    } catch (error) {
      console.error('Error saving temp data:', error);
      showSnackbar('Không thể lưu dữ liệu tạm. Vui lòng thử lại!', 'error');
    }
  };

  const loadTempFile = () => {
    try {
      const tempData = localStorage.getItem('tempData');
      if (tempData) {
        const parsedData = JSON.parse(tempData);
        
        // Restore data with attachments
        if (parsedData.data) {
          setData(parsedData.data);
        }
        
        return {
          inspector: parsedData.inspector,
          station: parsedData.station,
          date: parsedData.date
        };
      }
    } catch (error) {
      console.error('Error loading temp data:', error);
    }
    return null;
  };

  const loadTemplateFile = async () => {
    setLoading(true);
    try {
      const response = await fetch(process.env.PUBLIC_URL + '/data/template.xlsx');
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      // Lưu lại style và format của template
      const templateRange = worksheet['!ref'] || '';
      const templateMerges = worksheet['!merges'] || [];
      const templateCols = worksheet['!cols'] || [];

      // Chuyển đổi dữ liệu từ file Excel sang JSON
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1,
        raw: false,
        defval: null
      }) as any[];
      
      // Lấy header từ dòng đầu tiên
      const originalHeaders = jsonData[0] as string[];
      const powerAppsIdIndex = originalHeaders.findIndex(header => header === '__PowerAppsId__');
      let headerRow = originalHeaders.filter((header, index) => {
        return header && header !== '__PowerAppsId__';
      });
      
      // Thêm cột Date, INSPECTOR, STATION và attachment nếu chưa có
      if (!headerRow.includes('DATE')) {
        headerRow.push('DATE');
      }
      if (!headerRow.includes('Date')) {
        headerRow.push('Date');
      }
      if (!headerRow.includes('INSPECTOR')) {
        headerRow.push('INSPECTOR');
      }
      if (!headerRow.includes('STATION')) {
        headerRow.push('STATION');
      }
      if (!headerRow.includes('attachment')) {
        headerRow.push('attachment');
      }

      const orderedHeaders = ['STT', 'Date', 'STATION'];
      headerRow = [
        ...orderedHeaders,
        ...headerRow.filter(header => !orderedHeaders.includes(header))
      ];
      
      setHeaders(headerRow);

      // Tạo template row với các trường có thể chỉnh sửa
      const currentDate = new Date();
      const formattedDate = formatDate(currentDate);

      const templateRow: TemplateRow = {};
      headerRow.forEach(header => {
        templateRow[header] = {
          value: header === 'DATE' || header === 'Date' ? formattedDate : null,
          isEditable: ['INSPECTOR', 'STATION', 'DATE', 'Status', 'Note', 'Corrective action', 'Target', 'attachment'].includes(header)
        };
      });

      // Lọc dữ liệu, chỉ lấy các dòng có STT
      const filteredData = jsonData.slice(1)
        .filter((row: any[]) => row[originalHeaders.indexOf('STT')])
        .map((row: any[]) => {
          const rowData: { [key: string]: any } = {};
          headerRow.forEach(header => {
            if (header === 'Date' || header === 'DATE') {
              rowData[header] = formattedDate;
            } else if (header === 'INSPECTOR') {
              rowData[header] = '';
            } else if (header === 'STATION') {
              rowData[header] = '';
            } else if (header === 'attachment') {
              rowData[header] = [];
            } else {
              const originalIndex = originalHeaders.indexOf(header);
              if (originalIndex >= 0 && originalIndex !== powerAppsIdIndex) {
                rowData[header] = row[originalIndex] || '';
              }
            }
          });
          return rowData;
        });

      setData(filteredData);
      setTemplate(templateRow);

      // Lưu lại format của Excel để sử dụng khi xuất file
      setExcelFormat({
        range: templateRange,
        merges: templateMerges,
        cols: templateCols
      });
      
      // Reset các giá trị khi load template mới
      setSelectedInspector('');
      setSelectedStation('');
      setSelectedDate(currentDate);
      setSelectedSTT(1);

      const tempData = loadTempFile();
      if (tempData) {
        try {
          const { inspector, station, date } = tempData;
          if (inspector) {
            setSelectedInspector(inspector);
          }
          if (station) {
            setSelectedStation(station);
            // Cập nhật giá trị STATION cho tất cả các dòng
            setData(prev => prev.map(row => ({
              ...row,
              'STATION': station
            })));
          }
          if (date) {
            const parsedDate = parseDate(date);
            if (!isNaN(parsedDate.getTime())) {
              setSelectedDate(parsedDate);
              const formattedTempDate = formatDate(parsedDate);
              setData(prev => prev.map(row => ({
                ...row,
                'DATE': formattedTempDate,
                'Date': formattedTempDate
              })));
            }
          }
        } catch (error) {
          console.error('Error loading temp data:', error);
        }
      }
      setLoading(false);
      showSnackbar('Template loaded successfully!');
    } catch (error) {
      setLoading(false);
      console.error('Không thể tải file mẫu:', error);
      showSnackbar('Failed to load template!', 'error');
    }
  };

  const handleInputChange = (field: string, value: string | number | Date) => {
    if (template[field]?.isEditable) {
      try {
        let processedValue: string | number | null;
        
        // Xử lý đặc biệt cho trường Date
        if (field === 'Date' || field === 'DATE') {
          if (value instanceof Date) {
            processedValue = formatDate(value);
          } else if (typeof value === 'string') {
            const parsedDate = parseDate(value);
            processedValue = formatDate(parsedDate);
          } else {
            processedValue = null;
          }
        } else if (field === 'STT' && typeof value === 'string') {
          // Xử lý đặc biệt cho các trường số
          const numValue = parseInt(value);
          processedValue = !isNaN(numValue) ? numValue : null;
        } else {
          // Các trường khác
          processedValue = value as string | number;
        }
        
        setTemplate(prev => ({
          ...prev,
          [field]: {
            ...prev[field],
            value: processedValue
          }
        }));
        
        // Nếu trường là INSPECTOR, cập nhật selectedInspector
        if (field === 'INSPECTOR' && typeof value === 'string') {
          setSelectedInspector(value);
        }
        
        // Nếu trường là STATION, cập nhật selectedStation và tất cả các dòng
        if (field === 'STATION' && typeof value === 'string') {
          setSelectedStation(value);
          
          // Cập nhật giá trị STATION cho tất cả các dòng
          setData(prev => prev.map(row => ({
            ...row,
            'STATION': value
          })));
        }
        
        // Lưu dữ liệu tạm sau khi thay đổi
        saveTempFile();
      } catch (error) {
        console.error(`Error processing input for field ${field}:`, error);
        showSnackbar(`Lỗi khi cập nhật trường ${field}`, 'error');
      }
    }
  };

  const handleDateChange = (date: Date | null) => {
    if (date && !isNaN(date.getTime())) {
      const formattedDate = formatDate(date);
      
      // Cập nhật selectedDate
      setSelectedDate(date);
      
      // Cập nhật giá trị DATE và Date trong template
      setTemplate(prev => ({
        ...prev,
        'DATE': {
          ...prev['DATE'],
          value: formattedDate
        },
        'Date': {
          ...prev['Date'],
          value: formattedDate
        }
      }));

      // Cập nhật giá trị Date trong tất cả các dòng data
      setData(prev => prev.map(row => ({
        ...row,
        'DATE': formattedDate,
        'Date': formattedDate
      })));

      // Lưu dữ liệu tạm
      saveTempFile();
    }
  };

  const exportToExcel = async () => {
    setLoading(true);
    try {
      if (!selectedInspector) {
        showSnackbar('Vui lòng chọn INSPECTOR trước khi xuất file!', 'warning');
        setLoading(false);
        return;
      }

      if (!selectedStation) {
        showSnackbar('Vui lòng chọn STATION trước khi xuất file!', 'warning');
        setLoading(false);
        return;
      }

      // Kiểm tra xem có dữ liệu để xuất không
      const hasData = data.some(row => 
        Object.values(row).some(value => 
          value !== null && value !== undefined && value !== ''
        )
      );

      if (!hasData) {
        showSnackbar('Không có dữ liệu để xuất!', 'warning');
        setLoading(false);
        return;
      }

      // Lọc bỏ cột DATE khỏi headers khi xuất Excel
      const excelHeaders = headers.filter(header => header !== 'DATE');
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');

      // Add header row
      worksheet.addRow(excelHeaders);
      
      // Format header row
      const headerRow = worksheet.getRow(1);
      headerRow.font = { 
        bold: true,
        name: 'Arial',
        size: 12
      };
      
      // Áp dụng định dạng cho tất cả các ô trong header row
      headerRow.eachCell((cell) => {
        cell.alignment = { 
          vertical: 'middle', 
          horizontal: 'center',
          wrapText: true 
        };
      });
      
      // Set column widths
      excelHeaders.forEach((header, index) => {
        const column = worksheet.getColumn(index + 1);
        if (header === 'attachment') {
          column.width = 35; // Tăng độ rộng cột chứa ảnh
        } else {
          column.width = 15;
        }
      });

      // Add data rows
      const formattedDate = formatDate(selectedDate);
      
      data.forEach((row, rowIndex) => {
        const rowData = excelHeaders.map(header => {
          if (header === 'Date') {
            return formattedDate;
          }
          if (header === 'INSPECTOR') {
            return selectedInspector;
          }
          if (header === 'STATION') {
            return selectedStation;
          }
          if (header === 'attachment') {
            // Để ô trống thay vì hiển thị text
            return '';
          }
          return row[header] || '';
        });
        
        // Add the row data
        const excelRow = worksheet.addRow(rowData);
        
        // Tăng chiều cao của hàng để hiển thị nhiều ảnh
        excelRow.height = 200; // Tăng chiều cao để có thể hiển thị nhiều ảnh
        
        // Áp dụng định dạng cho tất cả các ô trong data row
        excelRow.eachCell((cell, colNumber) => {
          cell.font = {
            name: 'Arial',
            size: 12
          };
          
          const header = excelHeaders[colNumber - 1];
          
          // Special handling for Description column to preserve line breaks
          if (header === 'Description') {
            cell.alignment = {
              vertical: 'top',
              wrapText: true
            };
          } else {
            cell.alignment = {
              vertical: 'middle',
              wrapText: true
            };
          }
        });
        
        // Add images if available
        const attachments = row['attachment'] as ImageAttachment[] || [];
        if (attachments.length > 0) {
          const attachmentColIndex = excelHeaders.indexOf('attachment');
          if (attachmentColIndex !== -1) {
            // Xử lý tất cả các ảnh cho mỗi dòng
            const maxImagesPerRow = 4; // Giới hạn số lượng ảnh mỗi dòng để tránh quá tải
            const imagesToProcess = attachments.slice(0, maxImagesPerRow);
            
            // Tính toán kích thước và vị trí cho mỗi ảnh
            // Tăng kích thước ảnh thêm 15%
            const imageWidth = 92; // 80 + 15% = 92
            const imageHeight = 92; // 80 + 15% = 92
            
            // Sắp xếp ảnh theo lưới 2x2 nếu có nhiều ảnh
            imagesToProcess.forEach((attachment, imgIndex) => {
              try {
                const imageData = attachment.dataUrl;
                const base64Data = imageData.split(',')[1];
                
                // Thêm ảnh vào workbook
                const imageId = workbook.addImage({
                  base64: base64Data,
                  extension: 'jpeg',
                });
                
                // Tính toán vị trí dựa trên chỉ số ảnh với khoảng cách lớn hơn để tránh chồng lên nhau
                // Sắp xếp theo lưới 2x2: 0 1
                //                        2 3
                // Tăng khoảng cách giữa các ảnh
                const col = attachmentColIndex + (imgIndex % 2) * 0.6; // 0, 0.6, 0, 0.6
                const row = (rowIndex + 1) + Math.floor(imgIndex / 2) * 0.6; // Dòng + 0, 0, 0.6, 0.6
                
                // Thêm ảnh vào worksheet với vị trí đã tính
                worksheet.addImage(imageId, {
                  tl: { col: col, row: row },
                  ext: { width: imageWidth, height: imageHeight },
                  editAs: 'oneCell'
                });
              } catch (error) {
                console.error(`Error adding image ${imgIndex} to Excel:`, error);
              }
            });
            
            // Thêm chú thích về số lượng ảnh nếu có nhiều hơn giới hạn
            if (attachments.length > maxImagesPerRow) {
              const cell = worksheet.getCell(rowIndex + 2, attachmentColIndex + 1);
              cell.value = `+${attachments.length - maxImagesPerRow} more images`;
              cell.font = { name: 'Arial', size: 8, color: { argb: 'FF0000FF' } };
              cell.alignment = { vertical: 'middle', wrapText: true };
            }
          }
        }
      });
      
      // Đặt border cho tất cả các ô có dữ liệu
      const totalRows = worksheet.rowCount;
      const totalCols = worksheet.columnCount;
      
      for (let rowNum = 1; rowNum <= totalRows; rowNum++) {
        for (let colNum = 1; colNum <= totalCols; colNum++) {
          const cell = worksheet.getCell(rowNum, colNum);
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
        }
      }
      
      // Save workbook to file
      const dateStr = format(selectedDate, 'ddMMyyyy');
      const fileName = `Daily Nightshift report_${dateStr}_${selectedInspector}_${selectedStation}.xlsx`;
      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), fileName);
      
      showSnackbar('File đã được xuất thành công!', 'success');
      deleteTempFile();
      setLoading(false);
    } catch (error) {
      console.error('Error exporting file:', error);
      showSnackbar('Có lỗi khi xuất file. Vui lòng thử lại!', 'error');
      setLoading(false);
    }
  };

  const handleSubmit = () => {
    if (!selectedInspector) {
      alert('Vui lòng chọn INSPECTOR!');
      return;
    }

    if (!selectedStation) {
      alert('Vui lòng chọn STATION!');
      return;
    }

    // Tạo một bản sao của dữ liệu hiện tại
    const updatedData = [...data];
    const rowIndex = updatedData.findIndex(row => Number(row.STT) === selectedSTT);
    
    if (rowIndex === -1) {
      console.error('Không tìm thấy STT:', selectedSTT);
      return;
    }

    // Cập nhật dữ liệu cho dòng hiện tại
    const updatedRow = { ...updatedData[rowIndex] };
    headers.forEach(header => {
      if (header === 'INSPECTOR' && selectedInspector) {
        updatedRow[header] = selectedInspector;
      } else if ((header === 'Date' || header === 'DATE') && selectedDate) {
        updatedRow[header] = formatDate(selectedDate);
      } else if (header === 'STATION' && selectedStation) {
        updatedRow[header] = selectedStation;
      } else if (header === 'Status') {
        updatedRow[header] = template[header].value || 'Not Check';
      } else if (header === 'attachment') {
        // Preserve existing attachments
        updatedRow[header] = updatedRow[header] || [];
      } else if (template[header]?.isEditable) {
        updatedRow[header] = template[header].value;
      }
    });

    // Cập nhật dòng trong mảng dữ liệu
    updatedData[rowIndex] = updatedRow;
    
    // Cập nhật INSPECTOR và STATION cho tất cả các dòng
    const allUpdatedData = updatedData.map(row => ({
      ...row,
      INSPECTOR: selectedInspector,
      STATION: selectedStation
    }));
    
    setData(allUpdatedData);
    
    // Reset các trường có thể chỉnh sửa, ngoại trừ INSPECTOR, STATION và Date
    const resetTemplate = { ...template };
    headers.forEach(header => {
      if (resetTemplate[header]?.isEditable && 
          header !== 'INSPECTOR' && 
          header !== 'STATION' && 
          header !== 'Date' && 
          header !== 'DATE' &&
          header !== 'attachment') {
        resetTemplate[header].value = null;
      }
    });
    setTemplate(resetTemplate);
    
    // Tự động chuyển sang STT tiếp theo
    if (selectedSTT < data.length) {
      setSelectedSTT(selectedSTT + 1);
    }

    // Lưu dữ liệu tạm
    saveTempFile();
    showSnackbar('Data submitted successfully!');
  };

  const handleNewTemplate = () => {
    loadTemplateFile();
  };

  const handleSTTChange = (stt: number) => {
    if (stt >= 1 && stt <= data.length) {
      setSelectedSTT(stt);
      
      // Cập nhật template với dữ liệu của dòng được chọn
      const selectedRow = data.find(row => Number(row.STT) === stt);
      if (selectedRow) {
        const updatedTemplate = { ...template };
        headers.forEach(header => {
          if (updatedTemplate[header]) {
            if (header === 'attachment') {
              // Skip attachment field as it's not part of the template value
              return;
            }
            updatedTemplate[header] = {
              ...updatedTemplate[header],
              value: (typeof selectedRow[header] === 'string' || typeof selectedRow[header] === 'number') 
                ? selectedRow[header] as string | number | null 
                : null
            };
          }
        });
        setTemplate(updatedTemplate);
      }
    }
  };

  // Image handling functions
  const handleImageUpload = () => {
    if (fileInputRef.current) {
      fileInputRef.current.click();
    }
  };

  const resizeImage = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = (event) => {
        const img = new Image();
        img.src = event.target?.result as string;
        img.onload = () => {
          const canvas = document.createElement('canvas');
          let width = img.width;
          let height = img.height;
          
          // Calculate the new dimensions while maintaining aspect ratio
          if (width > MAX_IMAGE_WIDTH) {
            height = Math.round(height * (MAX_IMAGE_WIDTH / width));
            width = MAX_IMAGE_WIDTH;
          }
          
          if (height > MAX_IMAGE_HEIGHT) {
            width = Math.round(width * (MAX_IMAGE_HEIGHT / height));
            height = MAX_IMAGE_HEIGHT;
          }
          
          canvas.width = width;
          canvas.height = height;
          
          const ctx = canvas.getContext('2d');
          ctx?.drawImage(img, 0, 0, width, height);
          
          // Convert to data URL (JPEG format with 0.8 quality)
          const dataUrl = canvas.toDataURL('image/jpeg', 0.8);
          resolve(dataUrl);
        };
        img.onerror = () => {
          reject(new Error('Failed to load image'));
        };
      };
      reader.onerror = () => {
        reject(new Error('Failed to read file'));
      };
    });
  };

  const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (!files || files.length === 0) return;
    
    try {
      setLoading(true);
      
      const updatedData = [...data];
      const rowIndex = updatedData.findIndex(row => Number(row.STT) === selectedSTT);
      
      if (rowIndex === -1) {
        showSnackbar('Không tìm thấy dòng dữ liệu tương ứng!', 'error');
        setLoading(false);
        return;
      }
      
      // Initialize attachments array if it doesn't exist
      if (!updatedData[rowIndex]['attachment']) {
        updatedData[rowIndex]['attachment'] = [];
      }
      
      // Get current attachments
      const currentAttachments = updatedData[rowIndex]['attachment'] as ImageAttachment[] || [];
      
      // Process each file
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        try {
          const resizedDataUrl = await resizeImage(file);
          currentAttachments.push({
            dataUrl: resizedDataUrl,
            name: file.name
          });
        } catch (error) {
          console.error('Error processing image:', error);
          showSnackbar(`Lỗi khi xử lý ảnh ${file.name}`, 'error');
        }
      }
      
      // Update the row with new attachments
      updatedData[rowIndex]['attachment'] = currentAttachments;
      setData(updatedData);
      
      // Save temp data
      saveTempFile();
      showSnackbar('Đã tải lên ảnh thành công!', 'success');
    } catch (error) {
      console.error('Error uploading images:', error);
      showSnackbar('Lỗi khi tải lên ảnh!', 'error');
    } finally {
      setLoading(false);
      // Reset file input
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
    }
  };

  const handleDeleteImage = (rowIndex: number, imageIndex: number) => {
    const updatedData = [...data];
    const attachments = updatedData[rowIndex]['attachment'] as ImageAttachment[] || [];
    
    if (attachments.length > imageIndex) {
      attachments.splice(imageIndex, 1);
      updatedData[rowIndex]['attachment'] = attachments;
      setData(updatedData);
      saveTempFile();
      showSnackbar('Đã xóa ảnh thành công!', 'success');
    }
  };

  const handlePreviewImage = (imageUrl: string) => {
    setSelectedImage(imageUrl);
    setImagePreviewOpen(true);
  };

  const handleCloseImagePreview = () => {
    setImagePreviewOpen(false);
  };

  const renderAttachmentCell = (attachments: ImageAttachment[], rowIndex: number) => {
    return (
      <Box sx={{ 
        width: '100%', 
        maxWidth: { xs: '120px', sm: '150px' },
        minWidth: { xs: '100px', sm: '120px' }
      }}>
        {Array.isArray(attachments) && attachments.length > 0 ? (
          <ImageList sx={{ width: '100%', height: { xs: 100, sm: 120 } }} cols={2} rowHeight={50}>
            {attachments.map((img, imgIndex) => (
              <ImageListItem key={imgIndex}>
                <img
                  src={img.dataUrl}
                  alt={img.name}
                  style={{ width: '100%', height: '100%', objectFit: 'cover' }}
                />
                <Box sx={{ display: 'flex', justifyContent: 'space-between', mt: 0.5 }}>
                  <IconButton 
                    size="small" 
                    onClick={() => handlePreviewImage(img.dataUrl)}
                  >
                    <VisibilityIcon fontSize="small" />
                  </IconButton>
                  <IconButton 
                    size="small" 
                    color="error" 
                    onClick={() => handleDeleteImage(rowIndex, imgIndex)}
                  >
                    <DeleteIcon fontSize="small" />
                  </IconButton>
                </Box>
              </ImageListItem>
            ))}
          </ImageList>
        ) : (
          <Typography variant="caption" color="text.secondary">
            No images
          </Typography>
        )}
      </Box>
    );
  };

  return (
    <Container maxWidth="md" sx={{ mt: 2, mb: 4, px: { xs: 1, sm: 2, md: 3 } }}>
      <Box sx={{ display: 'flex', flexDirection: 'column', alignItems: 'center', mb: 3 }}>
        <img 
          src={`${process.env.PUBLIC_URL}/vietjet-logo.svg?v=${new Date().getTime()}`} 
          alt="Vietjet Air Logo" 
          style={{ 
            width: '200px', 
            marginBottom: '15px',
            maxHeight: '50px',
            objectFit: 'contain'
          }} 
        />
        <Typography 
          variant="h4" 
          component="h1" 
          gutterBottom 
          align="center" 
          sx={{ 
            color: '#e30613',
            fontSize: { xs: '1.5rem', sm: '2rem', md: '2.125rem' }
          }}
        >
          Daily Nightshift Report
        </Typography>
        
        <Paper sx={{ p: { xs: 1.5, sm: 2 }, mb: 2, width: '100%' }}>
          <Grid container spacing={{ xs: 1.5, sm: 2 }}>
            <Grid item xs={12} sm={6} md={4}>
              <FormControl fullWidth size="small">
                <InputLabel>STT</InputLabel>
                <Select
                  value={selectedSTT}
                  onChange={(e) => handleSTTChange(Number(e.target.value))}
                >
                  {data.map((row, index) => (
                    <MenuItem key={index} value={Number(row.STT || 0)}>
                      {String(row.STT)}
                    </MenuItem>
                  ))}
                </Select>
              </FormControl>
            </Grid>
            
            <Grid item xs={12} sm={6} md={4}>
              <FormControl fullWidth size="small">
                <InputLabel>Inspector</InputLabel>
                <Select
                  value={selectedInspector}
                  onChange={(e) => handleInputChange('INSPECTOR', e.target.value)}
                >
                  {INSPECTORS.map((inspector) => (
                    <MenuItem key={inspector} value={inspector}>
                      {inspector}
                    </MenuItem>
                  ))}
                </Select>
                <FormHelperText>Bắt buộc</FormHelperText>
              </FormControl>
            </Grid>
            
            <Grid item xs={12} sm={6} md={4}>
              <FormControl fullWidth size="small">
                <InputLabel>Station</InputLabel>
                <Select
                  value={selectedStation}
                  onChange={(e) => handleInputChange('STATION', e.target.value)}
                >
                  {STATIONS.map((station) => (
                    <MenuItem key={station} value={station}>
                      {station}
                    </MenuItem>
                  ))}
                </Select>
                <FormHelperText>Bắt buộc</FormHelperText>
              </FormControl>
            </Grid>
            
            <Grid item xs={12} sm={6} md={4}>
              <LocalizationProvider dateAdapter={AdapterDateFns}>
                <DatePicker
                  label="Date *"
                  value={selectedDate}
                  onChange={handleDateChange}
                  format="dd-MM-yyyy"
                  slotProps={{
                    textField: {
                      fullWidth: true,
                      size: "small",
                      error: !selectedDate,
                      helperText: !selectedDate ? 'Bắt buộc' : ''
                    }
                  }}
                />
              </LocalizationProvider>
            </Grid>

            <Grid item xs={12}>
              {selectedSTT && (
                <Paper elevation={0} sx={{ p: { xs: 1.5, sm: 2 }, bgcolor: '#f5f5f5', mb: { xs: 1.5, sm: 2 } }}>
                  <Typography variant="subtitle2" color="primary" gutterBottom>
                    Hướng dẫn kiểm tra:
                  </Typography>
                  <Typography 
                    variant="body2" 
                    sx={{ whiteSpace: 'pre-line', pl: 1, fontSize: { xs: '0.875rem', sm: '0.875rem' } }}
                  >
                    {getFieldInstruction(selectedSTT)}
                  </Typography>
                </Paper>
              )}
            </Grid>

            <Grid item xs={12}>
              <FormControl fullWidth>
                <InputLabel>Status</InputLabel>
                <Select
                  value={template['Status']?.value as string || ''}
                  onChange={(e) => handleInputChange('Status', e.target.value)}
                >
                  {STATUS_OPTIONS.map((option) => (
                    <MenuItem key={option} value={option}>
                      {option}
                    </MenuItem>
                  ))}
                </Select>
              </FormControl>
            </Grid>

            {headers
              .filter(header => !['STT', 'INSPECTOR', 'STATION', 'Status', 'Date', 'Note', 'Corrective action', 'Target', 'attachment'].includes(header))
              .map((header) => {
                const currentRow = data.find(row => Number(row.STT) === selectedSTT);
                const value = currentRow?.[header] || '';
                
                // Special handling for Description column to preserve whitespace and line breaks
                if (header === 'Description') {
                  return (
                    <Grid item xs={12} key={header}>
                      <TextField
                        fullWidth
                        label={header}
                        value={value}
                        InputProps={{
                          readOnly: true,
                          style: { whiteSpace: 'pre-wrap' }
                        }}
                        multiline
                        minRows={3}
                      />
                    </Grid>
                  );
                }
                
                return (
                  <Grid item xs={12} key={header}>
                    <TextField
                      fullWidth
                      label={header}
                      value={value}
                      InputProps={{
                        readOnly: true,
                      }}
                    />
                  </Grid>
                );
              })}

            <Grid item xs={12}>
              <TextField
                fullWidth
                label="Target"
                value={template['Target']?.value as string || ''}
                onChange={(e) => handleInputChange('Target', e.target.value)}
              />
            </Grid>

            <Grid item xs={12}>
              <TextField
                fullWidth
                label="Note"
                value={template['Note']?.value as string || ''}
                onChange={(e) => handleInputChange('Note', e.target.value)}
                multiline
                minRows={2}
              />
            </Grid>

            <Grid item xs={12}>
              <TextField
                fullWidth
                label="Corrective action"
                value={template['Corrective action']?.value as string || ''}
                onChange={(e) => handleInputChange('Corrective action', e.target.value)}
              />
            </Grid>
            
            <Grid item xs={12}>
              <Button
                fullWidth
                variant="contained"
                color="primary"
                onClick={handleImageUpload}
              >
                Tải ảnh
              </Button>
              <input
                type="file"
                multiple
                accept="image/*"
                style={{ display: 'none' }}
                ref={fileInputRef}
                onChange={handleFileChange}
              />
            </Grid>
          </Grid>
        </Paper>
        
        <Stack 
          direction={{ xs: 'column', sm: 'row' }} 
          spacing={{ xs: 1, sm: 2 }} 
          sx={{ mb: 2, width: '100%' }}
        >
          <Button
            fullWidth
            variant="contained"
            color="primary"
            onClick={handleSubmit}
            size="small"
          >
            Cập nhật dữ liệu
          </Button>
          <Button
            fullWidth
            variant="contained"
            color="secondary"
            onClick={exportToExcel}
            size="small"
          >
            Xuất File
          </Button>
        </Stack>
        
        <Stack 
          direction={{ xs: 'column', sm: 'row' }} 
          spacing={{ xs: 1, sm: 2 }}
          sx={{ width: '100%' }}
        >
          <Button
            fullWidth
            variant="outlined"
            onClick={saveTempFile}
            size="small"
          >
            Save Temp
          </Button>
          <Button
            fullWidth
            variant="outlined"
            color="error"
            onClick={deleteTempFile}
            size="small"
          >
            Delete Temp
          </Button>
        </Stack>
        
        {data.length > 0 && (
          <Box sx={{ mt: 4, width: '100%' }}>
            <Typography variant="h6" gutterBottom>
              Preview
            </Typography>
            <Box sx={{ overflowX: 'auto', maxWidth: '100%', WebkitOverflowScrolling: 'touch' }}>
              <table style={{ 
                width: '100%', 
                borderCollapse: 'collapse',
                minWidth: '650px' // Đảm bảo bảng có chiều rộng tối thiểu
              }}>
                <thead>
                  <tr>
                    {headers
                      .filter(header => header !== 'DATE')
                      .map((header) => (
                        <th key={header} style={{ 
                          padding: 8, 
                          borderBottom: '1px solid #ddd', 
                          textAlign: 'left',
                          whiteSpace: 'nowrap' // Ngăn header bị ngắt dòng
                        }}>
                          {header}
                        </th>
                      ))}
                  </tr>
                </thead>
                <tbody>
                  {data.map((row, rowIndex) => (
                    <tr key={rowIndex}>
                      {headers
                        .filter(header => header !== 'DATE')
                        .map((header) => (
                          <td key={header} style={{ 
                            padding: 8, 
                            borderBottom: '1px solid #ddd',
                            maxWidth: header === 'Description' ? '200px' : 
                                    header === 'attachment' ? '150px' : 
                                    header === 'Date' ? '100px' : 'auto',
                            minWidth: header === 'attachment' ? '120px' : 
                                     header === 'Date' ? '100px' : 'auto',
                            overflow: 'hidden',
                            textOverflow: 'ellipsis',
                            whiteSpace: header === 'Description' ? 'normal' : 'nowrap'
                          }}>
                            {header === 'Date' 
                              ? formatDate(selectedDate)
                              : header === 'attachment' 
                                ? renderAttachmentCell(row[header] as ImageAttachment[] || [], rowIndex)
                                : String(row[header] || '')}
                          </td>
                        ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </Box>
          </Box>
        )}
        
        <Snackbar
          open={snackbar.open}
          autoHideDuration={6000}
          onClose={handleCloseSnackbar}
        >
          <Alert
            severity={snackbar.severity}
            sx={{ width: '100%' }}
          >
            {snackbar.message}
          </Alert>
        </Snackbar>
        
        {loading && (
          <Box sx={{ position: 'fixed', top: 0, left: 0, width: '100%', height: '100%', display: 'flex', justifyContent: 'center', alignItems: 'center', backgroundColor: 'rgba(0, 0, 0, 0.5)' }}>
            <CircularProgress />
          </Box>
        )}
        
        <Dialog
          open={imagePreviewOpen}
          onClose={handleCloseImagePreview}
          maxWidth="lg"
          fullWidth
          sx={{
            '& .MuiDialog-paper': {
              margin: { xs: '16px', sm: '32px' },
              width: { xs: 'calc(100% - 32px)', sm: 'auto' },
              maxHeight: { xs: 'calc(100% - 32px)', sm: 'auto' }
            }
          }}
        >
          <DialogTitle>
            <Typography variant="h6" component="div">
              Preview Image
            </Typography>
          </DialogTitle>
          <DialogContent dividers>
            <Box 
              sx={{ 
                width: '100%', 
                height: '100%', 
                display: 'flex', 
                justifyContent: 'center',
                overflow: 'auto'
              }}
            >
              <img 
                src={selectedImage} 
                alt="Preview" 
                style={{ 
                  maxWidth: '100%', 
                  maxHeight: '70vh',
                  objectFit: 'contain'
                }} 
              />
            </Box>
          </DialogContent>
          <DialogActions>
            <Button onClick={handleCloseImagePreview} color="primary">
              Đóng
            </Button>
          </DialogActions>
        </Dialog>
      </Box>
    </Container>
  );
}

export default App;
