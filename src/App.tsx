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
    } catch (error) {
      console.error('Error deleting temp data:', error);
    }
  };

  const saveTempFile = () => {
    if (!selectedInspector) return;

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
    } catch (error) {
      console.error('Error saving temp data:', error);
      alert('Không thể lưu dữ liệu tạm. Vui lòng thử lại!');
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
        return header && header !== '__PowerAppsId__';
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
      worksheet.getRow(1).font = { bold: true };
      
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
        
        // Tăng chiều cao của hàng để hiển thị ảnh tốt hơn
        excelRow.height = 100;
        
        // Add images if available
        const attachments = row['attachment'] as ImageAttachment[] || [];
        if (attachments.length > 0) {
          const attachmentColIndex = excelHeaders.indexOf('attachment');
          if (attachmentColIndex !== -1) {
            // Process only the first image for each row
            const imageData = attachments[0].dataUrl;
            const base64Data = imageData.split(',')[1];
            
            // Add image to worksheet
            try {
              const imageId = workbook.addImage({
                base64: base64Data,
                extension: 'jpeg',
              });
              
              // Tăng kích thước ảnh thêm 15%
              const imageWidth = 92; // 80 * 1.15
              const imageHeight = 92; // 80 * 1.15
              
              // Position is 0-indexed for tl, but 1-indexed for row references
              worksheet.addImage(imageId, {
                tl: { col: attachmentColIndex, row: rowIndex + 1 },
                ext: { width: imageWidth, height: imageHeight },
                editAs: 'oneCell' // Đảm bảo ảnh được chèn vào trong ô
              });
            } catch (error) {
              console.error('Error adding image to Excel:', error);
            }
          }
        }
      });
      
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
    setData(updatedData);
    
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

  return (
    <Container maxWidth="md" sx={{ mt: 4, mb: 4 }}>
      <Box sx={{ display: 'flex', flexDirection: 'column', alignItems: 'center', mb: 4 }}>
        <img 
          src={`${process.env.PUBLIC_URL}/vietjet-logo.svg`} 
          alt="Vietjet Air Logo" 
          style={{ 
            width: '250px', 
            marginBottom: '20px' 
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

            {headers
              .filter(header => !['STT', 'INSPECTOR', 'Status', 'Date', 'Note', 'Corrective action', 'Target'].includes(header))
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
                    />
                  </Grid>
                );
              })}
            
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
                  <tr key={rowIndex}>
                    {headers
                      .filter(header => header !== 'DATE')
                      .map((header) => (
                        <td key={header} style={{ padding: 8, borderBottom: '1px solid #ddd' }}>
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
