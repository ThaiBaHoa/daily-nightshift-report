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
  [key: string]: string | number | null;
}

interface TemplateRow {
  [key: string]: {
    value: string | number | null;
    isEditable: boolean;
  };
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
  "Not Checked",
  "Finding"
];

function App() {
  const [data, setData] = useState<DataRow[]>([]);
  const [template, setTemplate] = useState<TemplateRow>({});
  const [headers, setHeaders] = useState<string[]>([]);
  const [selectedInspector, setSelectedInspector] = useState<string>('');
  const [selectedDate, setSelectedDate] = useState<Date | null>(new Date());
  const [selectedSTT, setSelectedSTT] = useState<number>(1);
  const [excelFormat, setExcelFormat] = useState<{
    range: string;
    merges: any[];
    cols: any[];
  } | null>(null);

  useEffect(() => {
    loadTemplateFile();
    // Khôi phục dữ liệu tạm nếu có
    try {
      const tempDataStr = localStorage.getItem('tempData');
      if (tempDataStr) {
        const tempData = JSON.parse(tempDataStr);
        setData(tempData.data);
        setSelectedInspector(tempData.inspector);
        setSelectedDate(new Date(tempData.date));
      }
    } catch (error) {
      console.error('Error loading temp data:', error);
    }
    return () => {
      deleteTempFile();
    };
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
        date: selectedDate
      };
      localStorage.setItem('tempData', JSON.stringify(tempData));
    } catch (error) {
      console.error('Error saving temp data:', error);
    }
  };

  const loadTemplateFile = async () => {
    try {
      const response = await fetch(process.env.PUBLIC_URL + '/data/template.xlsx');
      const arrayBuffer = await response.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      // Lưu lại style và format của template
      const templateRange = worksheet['!ref'] || 'A1';
      const templateMerges = worksheet['!merges'] || [];
      const templateCols = worksheet['!cols'] || [];
      
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        defval: null,
        raw: false
      }) as Record<string, string | number | null>[];
      
      if (jsonData.length > 0) {
        // Lọc bỏ cột __PowerAppsId__
        const filteredData = jsonData.map(row => {
          const newRow = { ...row };
          delete newRow.__PowerAppsId__;
          return newRow;
        });

        const headers = Object.keys(filteredData[0]).filter(h => h !== '__PowerAppsId__');
        setHeaders(headers);
        
        const templateRow: TemplateRow = {};
        headers.forEach(header => {
          const value = filteredData[0][header];
          templateRow[header] = {
            value: value,
            isEditable: value === null || header === 'INSPECTOR' || header === 'Status' || header === 'Date'
          };
        });

        setTemplate(templateRow);
        setData(filteredData);
        setExcelFormat({
          range: templateRange,
          merges: templateMerges,
          cols: templateCols
        });
        
        // Reset các giá trị khi load template mới
        setSelectedInspector('');
        setSelectedDate(new Date());
        setSelectedSTT(1);
      }
    } catch (error) {
      console.error('Không thể tải file mẫu:', error);
    }
  };

  const handleInputChange = (field: string, value: string | number) => {
    if (template[field]?.isEditable) {
      if (field === 'INSPECTOR') {
        setSelectedInspector(value as string);
      } else if (field === 'Date') {
        const date = new Date(value);
        const formattedDate = format(date, 'dd/MM/yyyy');
        setSelectedDate(date);
        // Cập nhật ngày cho tất cả các dòng
        setData(prev => prev.map(row => ({
          ...row,
          Date: formattedDate,
          DATE: formattedDate // Thêm cập nhật cho trường DATE nếu có
        })));
      }
      setTemplate(prev => ({
        ...prev,
        [field]: {
          ...prev[field],
          value: field === 'Date' ? format(new Date(value), 'dd/MM/yyyy') : value
        }
      }));
    }
  };

  const handleSubmit = () => {
    if (!selectedInspector) {
      alert('Vui lòng chọn INSPECTOR trước khi cập nhật dữ liệu!');
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
      } else if (header === 'Date' && selectedDate) {
        updatedRow[header] = format(selectedDate, 'dd/MM/yyyy');
      } else if (header === 'Status') {
        updatedRow[header] = template[header].value || 'Not Checked';
      } else if (template[header].isEditable) {
        updatedRow[header] = template[header].value;
      }
    });

    // Cập nhật dòng trong mảng dữ liệu
    updatedData[rowIndex] = updatedRow;
    setData(updatedData);
    
    // Reset các trường có thể chỉnh sửa, ngoại trừ INSPECTOR và Date
    const resetTemplate = { ...template };
    headers.forEach(header => {
      if (resetTemplate[header].isEditable && header !== 'INSPECTOR' && header !== 'Date') {
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
  };

  const exportData = () => {
    if (!selectedInspector) {
      alert('Vui lòng chọn INSPECTOR trước khi xuất file!');
      return;
    }

    try {
      const ws = XLSX.utils.json_to_sheet(data.map(row => ({
        ...row,
        Date: row.Date ? new Date(parse(row.Date, 'dd/MM/yyyy', new Date()).setHours(0, 0, 0, 0)) : null,
        DATE: row.DATE ? new Date(parse(row.DATE, 'dd/MM/yyyy', new Date()).setHours(0, 0, 0, 0)) : null
      })));
      
      if (excelFormat) {
        ws['!ref'] = excelFormat.range;
        ws['!merges'] = excelFormat.merges;
        ws['!cols'] = excelFormat.cols;
      }

      // Định dạng ngày tháng cho cột Date và DATE
      for (let row = 0; row < data.length; row++) {
        const dateCell = ws[XLSX.utils.encode_cell({ r: row + 1, c: headers.indexOf('Date') })];
        const DATE_Cell = ws[XLSX.utils.encode_cell({ r: row + 1, c: headers.indexOf('DATE') })];
        
        if (dateCell) {
          dateCell.t = 'd';  // Set cell type to date
          dateCell.z = 'dd/mm/yyyy';  // Set date format
        }
        if (DATE_Cell) {
          DATE_Cell.t = 'd';  // Set cell type to date
          DATE_Cell.z = 'dd/mm/yyyy';  // Set date format
        }
      }

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      
      const fileName = `Daily Nightshift report_${format(selectedDate || new Date(), 'ddMMyyyy')}.xlsx`;
      XLSX.writeFile(wb, fileName);
      alert('File đã được xuất thành công!\nBạn có thể tìm thấy file trong thư mục Downloads của thiết bị.');

      // Xóa dữ liệu tạm sau khi xuất thành công
      deleteTempFile();
    } catch (error) {
      console.error('Error exporting file:', error);
      alert('Có lỗi khi xuất file. Vui lòng thử lại!');
    }
  };

  const handleNewTemplate = () => {
    loadTemplateFile();
  };

  const handleSTTChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const stt = parseInt(event.target.value);
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
              value: selectedRow[header]
            };
          }
        });
        setTemplate(updatedTemplate);
      }
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
                <Grid item xs={12}>
                  <TextField
                    fullWidth
                    label="STT *"
                    type="number"
                    value={selectedSTT}
                    onChange={handleSTTChange}
                    inputProps={{ min: 1, max: data.length }}
                    helperText={`Nhập STT từ 1 đến ${data.length}`}
                  />
                </Grid>

                <Grid item xs={12}>
                  <FormControl fullWidth>
                    <InputLabel id="inspector-label">INSPECTOR *</InputLabel>
                    <Select
                      labelId="inspector-label"
                      value={selectedInspector}
                      label="INSPECTOR *"
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
                      onChange={(newValue) => {
                        if (newValue) {
                          // Chuyển đổi ngày về đầu ngày để tránh vấn đề múi giờ
                          const date = new Date(newValue);
                          date.setHours(0, 0, 0, 0);
                          handleInputChange('Date', date.toISOString());
                        }
                      }}
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

                {headers
                  .filter(header => !['STT', 'INSPECTOR', 'Status', 'Date', 'DATE'].includes(header))
                  .map((header) => {
                    const currentRow = data.find(row => Number(row.STT) === selectedSTT);
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
                  onClick={exportData}
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
                    <tr key={index} style={{ backgroundColor: row.STT === selectedSTT ? '#f5f5f5' : 'transparent' }}>
                      {headers
                        .filter(header => header !== 'DATE')
                        .map((header) => (
                          <td key={header} style={{ padding: 8, borderBottom: '1px solid #ddd' }}>
                            {header === 'Date' && row[header] 
                              ? format(new Date(row[header] as string), 'dd/MM/yyyy')
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
