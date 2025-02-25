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
  const [selectedStt, setSelectedStt] = useState<string>('');
  const [selectedInspector, setSelectedInspector] = useState<string>('');
  const [selectedDate, setSelectedDate] = useState<Date>(new Date());
  const [excelFormat, setExcelFormat] = useState<any>(null);

  const formatDate = (date: Date): string => {
    try {
      if (!(date instanceof Date) || isNaN(date.getTime())) {
        return format(new Date(), 'dd/MM/yyyy');
      }
      return format(date, 'dd/MM/yyyy');
    } catch (error) {
      console.error('Error formatting date:', error);
      return format(new Date(), 'dd/MM/yyyy');
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
      const response = await fetch(process.env.PUBLIC_URL + '/data/template.xlsx');
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      const templateRange = worksheet['!ref'] || '';
      const templateMerges = worksheet['!merges'] || [];
      const templateCols = worksheet['!cols'] || [];

      const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1,
        raw: false,
        defval: null
      }) as any[];
      
      const originalHeaders = jsonData[0] as string[];
      const powerAppsIdIndex = originalHeaders.findIndex(header => header === '__PowerAppsId__');
      let headerRow = originalHeaders.filter((header, index) => {
        return header && header !== '__PowerAppsId__';
      });
      
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

      const currentDate = new Date();
      const formattedDate = formatDate(currentDate);

      const templateRow: TemplateRow = {};
      headerRow.forEach(header => {
        templateRow[header] = {
          value: header === 'DATE' || header === 'Date' ? formattedDate : null,
          isEditable: ['INSPECTOR', 'DATE', 'Status', 'Note', 'Corrective action', 'Target'].includes(header)
        };
      });

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

      setTemplate(templateRow);
      setData(filteredData);
      setExcelFormat({
        range: templateRange,
        merges: templateMerges,
        cols: templateCols
      });
      
      setSelectedInspector('');
      setSelectedDate(currentDate);
      setSelectedStt('');

      const tempData = loadTempFile();
      if (tempData) {
        try {
          const { inspector, date } = tempData;
          if (inspector) {
            setSelectedInspector(inspector);
          }
          if (date) {
            const parsedDate = new Date(date);
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
    } catch (error) {
      console.error('Không thể tải file mẫu:', error);
    }
  };

  const handleInputChange = (field: string, value: string | number | Date) => {
    if (template[field]?.isEditable) {
      if (field === 'INSPECTOR') {
        setSelectedInspector(value as string);
        setTemplate(prev => ({
          ...prev,
          [field]: {
            ...prev[field],
            value: value as string
          }
        }));
      } else if (field === 'Status') {
        setTemplate(prev => ({
          ...prev,
          [field]: {
            ...prev[field],
            value: value as string
          }
        }));
      } else if (field === 'Target' || field === 'Note' || field === 'Corrective action') {
        setTemplate(prev => ({
          ...prev,
          [field]: {
            ...prev[field],
            value: value as string
          }
        }));
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
                      value={selectedStt}
                      label="STT"
                      onChange={handleSttChange}
                    >
                      {data.map((row) => (
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
                      onChange={(e: SelectChangeEvent) => handleInputChange('INSPECTOR', e.target.value)}
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
                          value={template[header]?.isEditable ? template[header]?.value || '' : currentRow?.[header] || ''}
                          onChange={(e) => handleInputChange(header, e.target.value)}
                          disabled={!template[header]?.isEditable}
                          helperText={template[header]?.isEditable ? 'Cần nhập' : 'Giá trị mặc định'}
                          required={template[header]?.isEditable}
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
