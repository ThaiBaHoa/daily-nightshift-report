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

type DataRow = {
  [key: string]: string;
};

type TemplateField = {
  value: string;
  type?: string;
};

type TemplateRow = {
  [key: string]: TemplateField;
};

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
    try {
      const tempData = {
        inspector: selectedInspector,
        date: selectedDate.toISOString()
      };
      localStorage.setItem('tempData', JSON.stringify(tempData));
    } catch (error) {
      console.error('Error saving temp data:', error);
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
    try {
      const response = await fetch(`${process.env.PUBLIC_URL}/template.xlsx`);
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      
      // Đọc dữ liệu với header từ dòng đầu tiên
      const jsonData = XLSX.utils.sheet_to_json(worksheet);
      
      // Lấy headers từ các key của dòng đầu tiên
      const headers = Object.keys(jsonData[0] || {}).filter(header => header !== '__EMPTY');
      setHeaders(headers);

      // Tạo template với các trường từ headers
      const templateRow: { [key: string]: { value: string; type?: string } } = {};
      headers.forEach(header => {
        templateRow[header] = {
          value: '',
          type: header === 'Status' ? 'select' : 'text'
        };
      });
      setTemplate(templateRow);

      // Chuyển đổi dữ liệu từ file Excel
      const rows = jsonData.map((row: any) => {
        const dataRow: { [key: string]: string } = {};
        headers.forEach(header => {
          if (header === 'Date' || header === 'DATE') {
            dataRow[header] = formatDate(selectedDate);
          } else {
            dataRow[header] = String(row[header] || '');
          }
        });
        return dataRow;
      });
      setData(rows);

      // Lưu định dạng Excel để sử dụng khi xuất file
      setExcelFormat(workbook);

    } catch (error) {
      console.error('Error loading template:', error);
      alert('Không thể tải file template. Vui lòng thử lại!');
    }
  };

  const handleInputChange = (field: string, value: string) => {
    if (template[field]) {
      setTemplate(prev => ({
        ...prev,
        [field]: {
          ...prev[field],
          value: value
        }
      }));
    }
  };

  const handleDateChange = (date: Date | null) => {
    if (date) {
      setSelectedDate(date);
      const formattedDate = formatDate(date);
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
    }
  };

  const handleInspectorChange = (event: SelectChangeEvent) => {
    const value = event.target.value;
    setSelectedInspector(value);
    if (template['INSPECTOR']) {
      setTemplate(prev => ({
        ...prev,
        'INSPECTOR': {
          ...prev['INSPECTOR'],
          value: value
        }
      }));
    }
  };

  const exportToExcel = () => {
    try {
      if (!selectedInspector) {
        alert('Vui lòng chọn INSPECTOR trước khi xuất file!');
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

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      
      const fileName = `Daily Nightshift report_${formattedDate.replace(/\//g, '')}.xlsx`;
      XLSX.writeFile(wb, fileName);
      alert('File đã được xuất thành công!\nBạn có thể tìm thấy file trong thư mục Downloads của thiết bị.');

      deleteTempFile();
    } catch (error) {
      console.error('Error exporting file:', error);
      alert('Có lỗi khi xuất file. Vui lòng thử lại!');
    }
  };

  const handleSubmit = () => {
    if (!selectedInspector) {
      alert('Vui lòng chọn INSPECTOR!');
      return;
    }

    if (!selectedStt) {
      alert('Vui lòng chọn STT!');
      return;
    }

    const updatedData = [...data];
    const rowIndex = updatedData.findIndex(row => row.STT === selectedStt);
    
    if (rowIndex === -1) {
      return;
    }

    // Cập nhật dữ liệu cho dòng được chọn
    headers.forEach(header => {
      if (template[header]) {
        updatedData[rowIndex][header] = template[header].value;
      }
    });

    // Cập nhật INSPECTOR và Date
    updatedData[rowIndex]['INSPECTOR'] = selectedInspector;
    updatedData[rowIndex]['Date'] = formatDate(selectedDate);
    if (updatedData[rowIndex]['DATE']) {
      updatedData[rowIndex]['DATE'] = formatDate(selectedDate);
    }

    setData(updatedData);

    // Reset form
    const resetTemplate = { ...template };
    headers.forEach(header => {
      if (resetTemplate[header]) {
        resetTemplate[header].value = '';
      }
    });
    setTemplate(resetTemplate);
    
    // Chọn dòng tiếp theo nếu có
    const nextRow = data.find(row => Number(row.STT) > Number(selectedStt));
    if (nextRow?.STT) {
      setSelectedStt(nextRow.STT.toString());
    }
    
    saveTempFile();
  };

  const handleNewTemplate = () => {
    loadTemplateFile();
  };

  const handleSttChange = (event: SelectChangeEvent) => {
    const stt = event.target.value;
    setSelectedStt(stt);
    
    // Tìm dòng tương ứng với STT được chọn
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
                    const currentRow = data.find(row => Number(row.STT) === Number(selectedStt));
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
                  onClick={exportToExcel}
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
