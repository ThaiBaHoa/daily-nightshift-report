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
  Stack
} from '@mui/material';
import { AdapterDateFns } from '@mui/x-date-pickers/AdapterDateFns';
import { LocalizationProvider, DatePicker } from '@mui/x-date-pickers';
import * as XLSX from 'xlsx';
import { format, parse } from 'date-fns';

interface DataRow {
  STT: string;
  INSPECTOR: string;
  Date: string;
  DATE: string;
  [key: string]: string;
}

interface TemplateField {
  value: string;
  type?: string;
}

interface TemplateRow {
  [key: string]: TemplateField;
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
  const [selectedStt, setSelectedStt] = useState<string>('');
  const [selectedInspector, setSelectedInspector] = useState<string>('');
  const [selectedDate, setSelectedDate] = useState<Date>(new Date());
  const [excelFormat, setExcelFormat] = useState<any>(null);

  const formatDate = (date: Date): string => {
    return format(date, 'dd/MM/yyyy');
  };

  useEffect(() => {
    loadTemplateFile();
    loadTempData();
  }, []);

  const deleteTempFile = () => {
    try {
      localStorage.removeItem('tempData');
    } catch (error) {
      console.error('Error deleting temp data:', error);
    }
  };

  const saveTempFile = () => {
    try {
      // Tạo một workbook mới từ data hiện tại
      const ws = XLSX.utils.json_to_sheet(data, { header: headers });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

      // Lưu file
      const wbout = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
      const blob = new Blob([wbout], { type: 'application/octet-stream' });
      
      // Lưu vào localStorage
      const reader = new FileReader();
      reader.onload = function(e) {
        if (e.target?.result) {
          localStorage.setItem('tempData', JSON.stringify({
            data: data,
            lastModified: new Date().toISOString()
          }));
        }
      };
      reader.readAsDataURL(blob);
    } catch (error) {
      console.error('Error saving temp file:', error);
    }
  };

  const loadTempData = () => {
    try {
      const tempDataStr = localStorage.getItem('tempData');
      if (tempDataStr) {
        const tempData = JSON.parse(tempDataStr);
        if (tempData.data && Array.isArray(tempData.data)) {
          setData(tempData.data);
          if (tempData.data.length > 0) {
            // Cập nhật template với dữ liệu từ dòng đầu tiên
            const firstRow = tempData.data[0];
            const updatedTemplate = { ...template };
            headers.forEach(header => {
              if (updatedTemplate[header]) {
                updatedTemplate[header] = {
                  ...updatedTemplate[header],
                  value: firstRow[header] || ''
                };
              }
            });
            setTemplate(updatedTemplate);
          }
        }
      }
    } catch (error) {
      console.error('Error loading temp data:', error);
    }
  };

  const loadTemplateFile = async () => {
    try {
      // Thử các đường dẫn khác nhau
      const possiblePaths = [
        `${process.env.PUBLIC_URL}/data/template.xlsx`,
        './data/template.xlsx',
        '/daily-nightshift-report/data/template.xlsx'
      ];

      let response;
      let successPath;

      for (const path of possiblePaths) {
        try {
          console.log('Trying path:', path);
          response = await fetch(path);
          if (response.ok) {
            successPath = path;
            console.log('Successfully loaded from:', path);
            break;
          }
        } catch (err) {
          console.log('Failed to load from:', path, err);
        }
      }

      if (!response || !response.ok) {
        throw new Error(`Could not load template from any path. Last attempted: ${successPath}`);
      }

      console.log('Response:', {
        status: response.status,
        statusText: response.statusText,
        ok: response.ok,
        url: response.url
      });

      const arrayBuffer = await response.arrayBuffer();
      console.log('Successfully loaded file, size:', arrayBuffer.byteLength);

      const workbook = XLSX.read(arrayBuffer);
      console.log('Workbook sheets:', workbook.SheetNames);

      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      
      // Đọc dữ liệu với header từ dòng đầu tiên
      const jsonData = XLSX.utils.sheet_to_json(worksheet);
      console.log('Parsed data:', jsonData);
      
      if (!Array.isArray(jsonData) || jsonData.length === 0) {
        throw new Error('No data found in template');
      }

      // Lấy headers từ các key của dòng đầu tiên
      const headers = Object.keys(jsonData[0] || {}).filter(header => header !== '__EMPTY');
      console.log('Headers:', headers);
      
      if (headers.length === 0) {
        throw new Error('No headers found in template');
      }

      setHeaders(headers);

      // Tạo template với các trường từ headers
      const templateRow: TemplateRow = {};
      headers.forEach(header => {
        templateRow[header] = {
          value: '',
          type: ['Type', 'TITLE', 'Description'].includes(header) ? 'readonly' : 
                header === 'Status' ? 'select' : 'text'
        };
      });
      setTemplate(templateRow);

      // Chuyển đổi dữ liệu từ file Excel
      const rows = jsonData.map((row: any) => {
        const dataRow: DataRow = {
          STT: String(row.STT || ''),
          INSPECTOR: String(row.INSPECTOR || ''),
          Date: formatDate(selectedDate),
          DATE: formatDate(selectedDate)
        };
        
        headers.forEach(header => {
          if (!['STT', 'INSPECTOR', 'Date', 'DATE'].includes(header)) {
            dataRow[header] = String(row[header] || '');
          }
        });
        
        return dataRow;
      });
      setData(rows);

      setExcelFormat(workbook);
      console.log('Template loaded successfully');
    } catch (error: any) {
      console.error('Detailed error loading template:', error);
      alert(`Không thể tải file template: ${error.message}`);
    }
  };

  const handleInputChange = (field: string, value: string) => {
    if (template[field] && selectedStt) {
      // Cập nhật template hiện tại
      setTemplate(prev => ({
        ...prev,
        [field]: {
          ...prev[field],
          value: value
        }
      }));

      // Cập nhật data tại dòng tương ứng với selectedStt
      setData(prev => prev.map(row => {
        if (row.STT === selectedStt) {
          return {
            ...row,
            [field]: value
          };
        }
        return row;
      }));
    }
  };

  const handleDateChange = (date: Date | null) => {
    if (date && selectedStt) {
      setSelectedDate(date);
      const formattedDate = formatDate(date);
      
      // Cập nhật template
      setTemplate(prev => {
        const newTemplate = { ...prev };
        if (newTemplate['Date']) {
          newTemplate['Date'].value = formattedDate;
        }
        if (newTemplate['DATE']) {
          newTemplate['DATE'].value = formattedDate;
        }
        return newTemplate;
      });

      // Cập nhật data tại dòng tương ứng với selectedStt
      setData(prev => prev.map(row => {
        if (row.STT === selectedStt) {
          return {
            ...row,
            'Date': formattedDate,
            'DATE': formattedDate
          };
        }
        return row;
      }));
    }
  };

  const handleInspectorChange = (event: SelectChangeEvent) => {
    const value = event.target.value;
    setSelectedInspector(value);
    if (template['INSPECTOR'] && selectedStt) {
      // Cập nhật template
      setTemplate(prev => ({
        ...prev,
        'INSPECTOR': {
          ...prev['INSPECTOR'],
          value: value
        }
      }));

      // Cập nhật data tại dòng tương ứng với selectedStt
      setData(prev => prev.map(row => {
        if (row.STT === selectedStt) {
          return {
            ...row,
            'INSPECTOR': value
          };
        }
        return row;
      }));
    }
  };

  const handleSttChange = (event: SelectChangeEvent) => {
    const stt = event.target.value;
    setSelectedStt(stt);
    
    // Tìm dòng tương ứng với STT được chọn
    const selectedRow = data.find(row => String(row.STT) === String(stt));
    if (selectedRow) {
      // Cập nhật template với dữ liệu từ dòng được chọn
      const updatedTemplate = { ...template };
      headers.forEach(header => {
        if (updatedTemplate[header]) {
          updatedTemplate[header] = {
            ...updatedTemplate[header],
            value: selectedRow[header] || ''
          };
        }
      });
      setTemplate(updatedTemplate);
    }
  };

  const handleExportExcel = () => {
    try {
      // Tạo một workbook mới từ data hiện tại
      const ws = XLSX.utils.json_to_sheet(data, { header: headers });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

      // Xuất file
      XLSX.writeFile(wb, `daily_report_${format(selectedDate, 'dd-MM-yyyy')}.xlsx`);
    } catch (error) {
      console.error('Error exporting file:', error);
      alert('Không thể xuất file. Vui lòng thử lại!');
    }
  };

  const handleSubmit = () => {
    if (!selectedStt) {
      alert('Vui lòng chọn STT!');
      return;
    }

    if (!selectedInspector) {
      alert('Vui lòng chọn INSPECTOR!');
      return;
    }

    // Tìm index của dòng cần cập nhật
    const rowIndex = data.findIndex(row => String(row.STT) === String(selectedStt));
    if (rowIndex === -1) {
      alert('Không tìm thấy dòng tương ứng với STT!');
      return;
    }

    // Tạo bản sao của data hiện tại
    const updatedData = [...data];
    
    // Cập nhật dữ liệu cho dòng được chọn
    const updatedRow: DataRow = { 
      ...updatedData[rowIndex],
      STT: selectedStt,
      INSPECTOR: selectedInspector,
      Date: formatDate(selectedDate),
      DATE: formatDate(selectedDate)
    };

    // Cập nhật các trường khác từ template
    headers.forEach(header => {
      if (template[header]) {
        updatedRow[header] = template[header].value;
      }
    });

    // Cập nhật dòng vào data
    updatedData[rowIndex] = updatedRow;
    setData(updatedData);

    // Lưu dữ liệu tạm
    saveTempFile();

    alert('Đã cập nhật dữ liệu thành công!');

    // Chọn dòng tiếp theo nếu có
    const nextRow = data.find(row => Number(row.STT) > Number(selectedStt));
    if (nextRow?.STT) {
      setSelectedStt(String(nextRow.STT));
    }
  };

  const handleNewTemplate = () => {
    loadTemplateFile();
  };

  return (
    <Container maxWidth="sm">
      <Box sx={{ my: 4 }}>
        <Typography variant="h4" component="h1" gutterBottom>
          Nhập dữ liệu Excel
        </Typography>
        
        <Paper sx={{ p: 2, mb: 2 }}>
          <Grid container spacing={2}>
            <Grid item xs={12} md={4}>
              <FormControl fullWidth>
                <InputLabel>STT</InputLabel>
                <Select
                  value={selectedStt}
                  onChange={(e) => {
                    const stt = e.target.value;
                    setSelectedStt(stt);
                    
                    // Tìm dòng dữ liệu tương ứng
                    const selectedRow = data.find(row => row.STT === stt);
                    if (selectedRow) {
                      // Cập nhật template với dữ liệu từ dòng được chọn
                      const updatedTemplate = { ...template };
                      headers.forEach(header => {
                        if (updatedTemplate[header]) {
                          updatedTemplate[header] = {
                            ...updatedTemplate[header],
                            value: selectedRow[header] || ''
                          };
                        }
                      });
                      setTemplate(updatedTemplate);
                    }
                  }}
                >
                  {data.map((row) => (
                    <MenuItem key={row.STT} value={row.STT}>
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
                  onChange={(e) => {
                    const inspector = e.target.value;
                    setSelectedInspector(inspector);
                    handleInputChange('INSPECTOR', inspector);
                  }}
                >
                  {INSPECTORS.map((inspector) => (
                    <MenuItem key={inspector} value={inspector}>
                      {inspector}
                    </MenuItem>
                  ))}
                </Select>
              </FormControl>
            </Grid>

            <Grid item xs={12} md={4}>
              <LocalizationProvider dateAdapter={AdapterDateFns}>
                <DatePicker
                  label="Date"
                  value={selectedDate}
                  onChange={handleDateChange}
                  format="dd/MM/yyyy"
                />
              </LocalizationProvider>
            </Grid>

            {headers.map((header) => {
              if (!['STT', 'INSPECTOR', 'Date', 'DATE'].includes(header) && template[header]) {
                if (template[header].type === 'select') {
                  return (
                    <Grid item xs={12} md={4} key={header}>
                      <FormControl fullWidth>
                        <InputLabel>{header}</InputLabel>
                        <Select
                          value={template[header].value}
                          onChange={(e) => handleInputChange(header, e.target.value)}
                        >
                          {STATUS_OPTIONS.map((option) => (
                            <MenuItem key={option} value={option}>
                              {option}
                            </MenuItem>
                          ))}
                        </Select>
                      </FormControl>
                    </Grid>
                  );
                } else if (template[header].type === 'readonly') {
                  return (
                    <Grid item xs={12} md={4} key={header}>
                      <TextField
                        fullWidth
                        label={header}
                        value={template[header].value}
                        InputProps={{
                          readOnly: true,
                        }}
                      />
                    </Grid>
                  );
                } else {
                  return (
                    <Grid item xs={12} md={4} key={header}>
                      <TextField
                        fullWidth
                        label={header}
                        value={template[header].value}
                        onChange={(e) => handleInputChange(header, e.target.value)}
                      />
                    </Grid>
                  );
                }
              }
              return null;
            })}
          </Grid>
        </Paper>

        <Stack direction="row" spacing={2} justifyContent="flex-end">
          <Button variant="contained" onClick={saveTempFile}>
            Save Temp
          </Button>
          <Button variant="contained" onClick={deleteTempFile}>
            Delete Temp
          </Button>
        </Stack>

        <Stack direction="row" spacing={2} justifyContent="flex-end" sx={{ mt: 2 }}>
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
            onClick={handleExportExcel}
          >
            Xuất File
          </Button>
        </Stack>
      </Box>
    </Container>
  );
}

export default App;
