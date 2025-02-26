import React, { useState, useEffect } from 'react';
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
  CircularProgress
} from '@mui/material';
import { AdapterDateFns } from '@mui/x-date-pickers/AdapterDateFns';
import { LocalizationProvider, DatePicker } from '@mui/x-date-pickers';
import * as XLSX from 'xlsx';
import { format, parse } from 'date-fns';

interface TemplateField {
  value: string | number | null;
  isEditable: boolean;
}

interface TemplateRow {
  [key: string]: TemplateField;
}

interface DataRow {
  [key: string]: string | number | null;
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
      }
      
      localStorage.setItem('tempData', tempDataString);
      console.log('Đã lưu dữ liệu tạm thời thành công');
    } catch (error) {
      console.error('Error saving temp data:', error);
      alert('Không thể lưu dữ liệu tạm. Vui lòng thử lại!');
    }
  };

  const loadTempFile = (): { inspector: string; date: string } | null => {
    try {
      const tempData = localStorage.getItem('tempData');
      if (tempData) {
        return JSON.parse(tempData);
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
      
      // Thêm cột Date và INSPECTOR nếu chưa có
      if (!headerRow.includes('DATE')) {
        headerRow.push('DATE');
      }
      if (!headerRow.includes('Date')) {
        headerRow.push('Date');
      }
      if (!headerRow.includes('INSPECTOR')) {
        headerRow.push('INSPECTOR');
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
          isEditable: ['INSPECTOR', 'DATE', 'Status', 'Note', 'Corrective action', 'Target'].includes(header)
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

  const exportToExcel = () => {
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
      const ws = XLSX.utils.aoa_to_sheet([excelHeaders]);

      const formattedDate = formatDate(selectedDate);
      const rows = data.map(row => {
        return excelHeaders.map(header => {
          if (header === 'Date') {
            return formattedDate;
          }
          if (header === 'INSPECTOR') {
            return selectedInspector;
          }
          return row[header] || '';
        });
      });

      XLSX.utils.sheet_add_aoa(ws, rows, { origin: -1 });

      // Thêm style cho worksheet
      ws['!cols'] = excelHeaders.map(() => ({ wch: 15 })); // Đặt chiều rộng cột
      
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      
      const fileName = `Daily Nightshift report_${formattedDate.replace(/\//g, '')}_${selectedInspector}.xlsx`;
      XLSX.writeFile(wb, fileName);
      
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
          header !== 'DATE') {
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
            updatedTemplate[header] = {
              ...updatedTemplate[header],
              value: selectedRow[header] || null
            };
          }
        });
        setTemplate(updatedTemplate);
      }
    }
  };

  return (
    <Container maxWidth="md">
      <Box sx={{ my: 4 }}>
        <Typography variant="h4" component="h1" gutterBottom align="center">
          Nhập dữ liệu Excel
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
                      {row.STT}
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
                {data.map((row, index) => (
                  <tr key={index}>
                    {headers
                      .filter(header => header !== 'DATE')
                      .map((header) => (
                        <td key={header} style={{ padding: 8, borderBottom: '1px solid #ddd' }}>
                          {header === 'Date' 
                            ? formatDate(selectedDate)
                            : row[header] || ''}
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
      </Box>
    </Container>
  );
}

export default App;
