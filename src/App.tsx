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
      // Thử tải từ public folder
      let response = await fetch(`${process.env.PUBLIC_URL}/template.xlsx`);
      
      // Nếu không tìm thấy trong public folder, thử tải từ thư mục gốc
      if (!response.ok) {
        response = await fetch('/template.xlsx');
      }

      // Nếu vẫn không tìm thấy, thử tải từ thư mục daily-nightshift-report
      if (!response.ok) {
        response = await fetch('/daily-nightshift-report/template.xlsx');
      }

      if (!response.ok) {
        throw new Error('Template file not found');
      }

      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      
      // Đọc dữ liệu với header từ dòng đầu tiên
      const jsonData = XLSX.utils.sheet_to_json(worksheet);
      
      // Lấy headers từ các key của dòng đầu tiên
      const headers = Object.keys(jsonData[0] || {}).filter(header => header !== '__EMPTY');
      setHeaders(headers);

      // Tạo template với các trường từ headers
      const templateRow: TemplateRow = {};
      headers.forEach(header => {
        templateRow[header] = {
          value: '',
          type: header === 'Status' ? 'select' : 'text'
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
    } catch (error) {
      console.error('Error loading template:', error);
      alert('Không thể tải file template. Vui lòng thử lại!');
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
          <Stack direction="row" spacing={2} sx={{ mb: 2 }}>
            <Button
              variant="contained"
              onClick={handleNewTemplate}
            >
              Mở Template Mới
            </Button>
          </Stack>

          {headers.length > 0 && (
            <Box component="form" noValidate sx={{ mt: 2 }}>
              <Grid container spacing={2}>
                <Grid item xs={12} sm={4}>
                  <FormControl fullWidth>
                    <InputLabel id="stt-select-label">STT</InputLabel>
                    <Select
                      labelId="stt-select-label"
                      id="stt-select"
                      value={selectedStt || ''}
                      label="STT"
                      onChange={handleSttChange}
                    >
                      {data.map((row) => row.STT && (
                        <MenuItem key={row.STT} value={row.STT}>
                          {row.STT}
                        </MenuItem>
                      ))}
                    </Select>
                  </FormControl>
                </Grid>

                <Grid item xs={12} sm={4}>
                  <FormControl fullWidth>
                    <InputLabel id="inspector-select-label">INSPECTOR</InputLabel>
                    <Select
                      labelId="inspector-select-label"
                      id="inspector-select"
                      value={selectedInspector}
                      label="INSPECTOR"
                      onChange={handleInspectorChange}
                      required
                    >
                      {INSPECTORS.map((inspector) => (
                        <MenuItem key={inspector} value={inspector}>
                          {inspector}
                        </MenuItem>
                      ))}
                    </Select>
                    <FormHelperText>Chọn người kiểm tra</FormHelperText>
                  </FormControl>
                </Grid>

                <Grid item xs={12}>
                  <FormControl fullWidth>
                    <InputLabel id="status-label">Status *</InputLabel>
                    <Select
                      labelId="status-label"
                      value={(template['Status']?.value || '').toString()}
                      label="Status *"
                      onChange={(e: SelectChangeEvent) => handleInputChange('Status', e.target.value)}
                      required
                    >
                      {STATUS_OPTIONS.map((status) => (
                        <MenuItem key={status} value={status}>
                          {status}
                        </MenuItem>
                      ))}
                    </Select>
                    <FormHelperText>Chọn trạng thái kiểm tra</FormHelperText>
                  </FormControl>
                </Grid>

                <Grid item xs={12}>
                  <LocalizationProvider dateAdapter={AdapterDateFns}>
                    <DatePicker
                      label="Date *"
                      value={selectedDate}
                      onChange={handleDateChange}
                      format="dd/MM/yyyy"
                      slotProps={{
                        textField: {
                          fullWidth: true,
                          required: true,
                          helperText: 'Chọn ngày kiểm tra (áp dụng cho tất cả các dòng)',
                          sx: { width: '100%' }
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
                    const currentRow = data.find(row => String(row.STT) === String(selectedStt));
                    return (
                      <Grid item xs={12} key={header}>
                        <TextField
                          fullWidth
                          label={header}
                          value={template[header]?.value || ''}
                          onChange={(e) => handleInputChange(header, e.target.value)}
                          helperText="Cần nhập"
                          required
                        />
                      </Grid>
                    );
                  })}
              </Grid>
              
              <Stack direction="row" spacing={2} sx={{ mt: 2 }}>
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
          )}
        </Paper>

        {data.length > 0 && (
          <Paper sx={{ p: 2 }}>
            <Typography variant="h6" gutterBottom>
              Dữ liệu hiện tại
            </Typography>
            <Box sx={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                <thead>
                  <tr>
                    {headers
                      .filter(header => header !== 'DATE')
                      .map((header) => (
                        <th key={header} style={{ padding: 8, borderBottom: '1px solid #ddd' }}>
                          {header}
                        </th>
                      ))}
                  </tr>
                </thead>
                <tbody>
                  {data.map((row, index) => (
                    <tr key={index} style={{ backgroundColor: row.STT === selectedStt ? '#f5f5f5' : 'transparent' }}>
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
          </Paper>
        )}
      </Box>
    </Container>
  );
}

export default App;
