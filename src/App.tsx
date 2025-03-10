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
import { FIELD_INSTRUCTIONS } from './fieldInstructions';

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
        return format(new Date(), 'dd/MM/yyyy');
      }
      return format(date, 'dd/MM/yyyy');
    } catch (error) {
      console.error('Error formatting date:', error);
      return format(new Date(), 'dd/MM/yyyy');
    }
  };

  const parseDate = (dateString: string): Date => {
    try {
      // Hỗ trợ nhiều định dạng ngày tháng phổ biến
      const formats = ['dd/MM/yyyy', 'yyyy-MM-dd', 'MM/dd/yyyy'];
      
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

    try {
      // Lưu dữ liệu vào localStorage thay vì file tạm
      const tempData = {
        data,
        inspector: selectedInspector,
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
        // Chỉ lọc bỏ __PowerAppsId__, giữ lại tất cả các header khác kể cả null hoặc empty
        return header !== '__PowerAppsId__';
      });
      
      // Thêm cột Date, INSPECTOR và attachment nếu chưa có
      if (!headerRow.includes('DATE')) {
        headerRow.push('DATE');
      }
      if (!headerRow.includes('Date')) {
        headerRow.push('Date');
      }
      if (!headerRow.includes('INSPECTOR')) {
        headerRow.push('INSPECTOR');
      }
      if (!headerRow.includes('attachment')) {
        headerRow.push('attachment');
      }

      const orderedHeaders = ['STT', 'Date'];
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
          isEditable: ['INSPECTOR', 'DATE', 'Status', 'Note', 'Corrective action', 'Target', 'attachment'].includes(header)
        };
      });

      // Lọc dữ liệu, lấy tất cả các dòng từ file Excel
      const filteredData = jsonData.slice(1)
        .map((row: any[]) => {
          const rowData: { [key: string]: any } = {};
          headerRow.forEach(header => {
            if (header === 'Date' || header === 'DATE') {
              rowData[header] = formattedDate;
            } else if (header === 'INSPECTOR') {
              rowData[header] = '';
            } else if (header === 'attachment') {
              rowData[header] = [];
            } else if (header === 'Description') {
              // Ensure Description field gets the full content
              const originalIndex = originalHeaders.indexOf(header);
              if (originalIndex >= 0 && originalIndex !== powerAppsIdIndex) {
                // Convert to string to ensure all content is preserved
                rowData[header] = String(row[originalIndex] || '');
              }
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
      setSelectedDate(currentDate);
      setSelectedSTT(1);

      const tempData = loadTempFile();
      if (tempData) {
        try {
          const { inspector, date } = tempData;
          if (inspector) {
            setSelectedInspector(inspector);
          }
          if (date) {
            const parsedDate = parseDate(date);
            if (!isNaN(parsedDate.getTime())) {
              setSelectedDate(parsedDate);
              const formattedTempDate = formatDate(parsedDate);
              setData(prev => prev.map(row => ({
                ...row,
                Date: formattedTempDate
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
          
          // Áp dụng định dạng đặc biệt cho cột Description
          const currentHeader = excelHeaders[colNumber - 1];
          if (currentHeader === 'Description') {
            cell.alignment = {
              vertical: 'top',
              horizontal: 'left',
              wrapText: true
            };
            // Tăng chiều cao của ô để hiển thị nhiều dòng văn bản
            if (cell.value && cell.value.toString().length > 50) {
              excelRow.height = Math.max(excelRow.height || 20, 100);
            }
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
      const fileName = `Daily Nightshift report_${formattedDate.replace(/\//g, '')}_${selectedInspector}.xlsx`;
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
    
    // Cập nhật INSPECTOR cho tất cả các dòng
    const allUpdatedData = updatedData.map(row => ({
      ...row,
      INSPECTOR: selectedInspector
    }));
    
    setData(allUpdatedData);
    
    // Reset các trường có thể chỉnh sửa, ngoại trừ INSPECTOR và Date
    const resetTemplate = { ...template };
    headers.forEach(header => {
      if (resetTemplate[header]?.isEditable && 
          header !== 'INSPECTOR' && 
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
              value: (typeof selectedRow[header] === 'string' || typeof selectedRow[header] === 'number' || selectedRow[header] === null) 
                ? selectedRow[header] as string | number | null 
                : null
            };
          }
        });
        setTemplate(updatedTemplate);
      }
    }
  };

  // Hàm để lấy hướng dẫn cho STT hiện tại
  const getCurrentInstruction = (): string | null => {
    const currentRow = data.find(row => Number(row.STT) === selectedSTT);
    if (!currentRow || !currentRow.STT) return null;
    
    return FIELD_INSTRUCTIONS[String(currentRow.STT)] || null;
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
      <Box>
        {Array.isArray(attachments) && attachments.length > 0 ? (
          <ImageList sx={{ width: 120, height: 120 }} cols={2} rowHeight={60}>
            {attachments.map((img, imgIndex) => (
              <ImageListItem key={imgIndex}>
                <img
                  src={img.dataUrl}
                  alt={img.name}
                  style={{ width: 50, height: 50, objectFit: 'cover' }}
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

  const renderFieldValue = (header: string, value: any, rowIndex: number) => {
    if (header === 'attachment' && Array.isArray(value)) {
      return renderAttachmentCell(value, rowIndex);
    } else if (header === 'Description') {
      return (
        <Box sx={{ 
          whiteSpace: 'pre-wrap', 
          wordBreak: 'break-word',
          minHeight: '80px',
          maxHeight: '200px',
          overflow: 'auto',
          border: '1px solid #e0e0e0',
          borderRadius: '4px',
          p: 1,
          bgcolor: '#f5f5f5',
          textAlign: 'left',
          fontSize: '0.875rem',
          lineHeight: '1.5'
        }}>
          {value}
        </Box>
      );
    } else if (header === 'Date') {
      return formatDate(selectedDate);
    } else {
      return String(value || '');
    }
  };

  return (
    <Container maxWidth="md" sx={{ mt: 4, mb: 4 }}>
      <Box sx={{ display: 'flex', flexDirection: 'column', alignItems: 'center', mb: 4 }}>
        <img 
          src={`${process.env.PUBLIC_URL}/vietjet-logo.svg?v=${new Date().getTime()}`} 
          alt="Vietjet Air Logo" 
          style={{ 
            width: '250px', 
            marginBottom: '20px',
            maxHeight: '60px',
            objectFit: 'contain'
          }} 
        />
        <Typography variant="h4" component="h1" gutterBottom align="center" sx={{ color: '#e30613' }}>
          Daily Nightshift Report
        </Typography>
        
        <Paper sx={{ p: 2, mb: 2 }}>
          <Grid container spacing={2}>
            <Grid item xs={12} md={4}>
              <FormControl fullWidth>
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
            
            <Grid item xs={12} md={4}>
              <FormControl fullWidth>
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
            
            <Grid item xs={12} md={4}>
              <LocalizationProvider dateAdapter={AdapterDateFns}>
                <DatePicker
                  label="Date *"
                  value={selectedDate}
                  onChange={handleDateChange}
                  format="dd/MM/yyyy"
                  slotProps={{
                    textField: {
                      fullWidth: true,
                      error: !selectedDate,
                      helperText: !selectedDate ? 'Bắt buộc' : ''
                    }
                  }}
                />
              </LocalizationProvider>
            </Grid>

            {/* Hiển thị hướng dẫn cho STT hiện tại nếu có */}
            {getCurrentInstruction() && (
              <Grid item xs={12}>
                <Paper elevation={0} sx={{ p: 2, bgcolor: '#f8f9fa', mb: 2, border: '1px solid #e0e0e0' }}>
                  <Typography variant="subtitle2" color="primary" gutterBottom>
                    Hướng dẫn kiểm tra:
                  </Typography>
                  <Typography variant="body2" style={{ whiteSpace: 'pre-line' }}>
                    {getCurrentInstruction()}
                  </Typography>
                </Paper>
              </Grid>
            )}

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
              .filter(header => !['STT', 'INSPECTOR', 'Status', 'Date', 'DATE', 'Note', 'Corrective action', 'Target', 'attachment'].includes(header))
              .map((header) => {
                const currentRow = data.find(row => Number(row.STT) === selectedSTT);
                return (
                  <Grid item xs={12} key={header}>
                    <TextField
                      fullWidth
                      label={header}
                      value={currentRow?.[header] || ''}
                      InputProps={{
                        readOnly: true,
                      }}
                      multiline={header === 'Description'}
                      minRows={header === 'Description' ? 3 : 1}
                      maxRows={header === 'Description' ? 10 : 1}
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
        
        <Stack direction="row" spacing={2} sx={{ mb: 2 }}>
          <Button
            fullWidth
            variant="contained"
            color="primary"
            onClick={handleSubmit}
          >
            Cập nhật dữ liệu
          </Button>
          <Button
            fullWidth
            variant="contained"
            color="secondary"
            onClick={exportToExcel}
          >
            Xuất File
          </Button>
        </Stack>
        
        <Stack direction="row" spacing={2}>
          <Button
            fullWidth
            variant="outlined"
            onClick={saveTempFile}
          >
            Save Temp
          </Button>
          <Button
            fullWidth
            variant="outlined"
            color="error"
            onClick={deleteTempFile}
          >
            Delete Temp
          </Button>
        </Stack>
        
        {data.length > 0 && (
          <Box sx={{ mt: 4, overflowX: 'auto' }}>
            <Typography variant="h6" gutterBottom>
              Preview
            </Typography>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  {headers
                    .filter(header => header !== 'DATE')
                    .map((header) => (
                      <th key={header} style={{ padding: 8, borderBottom: '1px solid #ddd', textAlign: 'left' }}>
                        {header}
                      </th>
                    ))}
                </tr>
              </thead>
              <tbody>
                {data.map((row, rowIndex) => (
                  <tr 
                    key={rowIndex} 
                    style={{ 
                      backgroundColor: Number(row.STT) === selectedSTT ? '#e3f2fd' : 'inherit',
                      cursor: 'pointer'
                    }}
                    onClick={() => handleSTTChange(Number(row.STT))}
                  >
                    {headers
                      .filter(header => header !== 'DATE')
                      .map((header, colIndex) => (
                        <td 
                          key={colIndex} 
                          style={{ 
                            padding: '8px', 
                            border: '1px solid #ddd',
                            textAlign: header === 'Description' ? 'left' : 'center',
                            whiteSpace: header === 'Description' ? 'pre-wrap' : 'normal',
                            maxWidth: header === 'Description' ? '300px' : 'auto',
                            minHeight: header === 'Description' ? '100px' : 'auto',
                            verticalAlign: header === 'Description' ? 'top' : 'middle'
                          }}>
                            {renderFieldValue(header, row[header], rowIndex)}
                          </td>
                      ))}
                  </tr>
                ))}
              </tbody>
            </table>
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
        >
          <DialogTitle>Preview Image</DialogTitle>
          <DialogContent>
            <img src={selectedImage} alt="Preview" style={{ width: '100%', height: '100%' }} />
          </DialogContent>
          <DialogActions>
            <Button onClick={handleCloseImagePreview} color="primary">
              Close
            </Button>
          </DialogActions>
        </Dialog>
      </Box>
    </Container>
  );
}

export default App;
