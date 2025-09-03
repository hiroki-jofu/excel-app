import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { styled } from '@mui/material/styles';
import {
  Button,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Paper,
  Container,
  Typography,
  Box,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  Grid,
  Card,
  CardContent,
  Divider
} from '@mui/material';
import UploadFileIcon from '@mui/icons-material/UploadFile';
import NoteAddIcon from '@mui/icons-material/NoteAdd';
import DownloadIcon from '@mui/icons-material/Download';

const StyledTableRow = styled(TableRow)(({ theme }) => ({
  '&:nth-of-type(odd)': {
    backgroundColor: theme.palette.action.hover,
  },
  // hide last border
  '&:last-child td, &:last-child th': {
    border: 0,
  },
}));

function App() {
  const [data, setData] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [sheetNames, setSheetNames] = useState([]);
  const [workbook, setWorkbook] = useState(null);
  const [selectedSheet, setSelectedSheet] = useState('');

  const processWorkbook = (wb) => {
    setWorkbook(wb);
    setSheetNames(wb.SheetNames);
    if (wb.SheetNames.length > 0) {
      handleSheetSelect(wb.SheetNames[0], wb);
    }
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (evt) => {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        processWorkbook(wb);
      };
      reader.readAsBinaryString(file);
    }
  };

  const handleLoadReferenceFile = () => {
    fetch('/データ変換例.xlsx')
      .then(res => res.arrayBuffer())
      .then(ab => {
        const wb = XLSX.read(ab, { type: 'buffer' });
        processWorkbook(wb);
      });
  };

  const handleSheetSelect = (sheetName, wb) => {
    setSelectedSheet(sheetName);
    const worksheet = (wb || workbook).Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    if (jsonData.length > 0) {
      setHeaders(jsonData[0]);
      setData(jsonData.slice(1));
    } else {
      setHeaders([]);
      setData([]);
    }
  };

  const handleExport = (format) => {
    const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
    if (format === 'csv') {
      const csv = XLSX.utils.sheet_to_csv(ws);
      const blob = new Blob(["\uFEFF" + csv], { type: 'text/csv;charset=utf-8;' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'data.csv';
      link.click();
    } else {
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
      XLSX.writeFile(wb, 'data.xlsx');
    }
  };

  const transformAtoAPrime = () => {
    const baseHeaders = headers.slice(0, 7);
    const newHeaders = [...baseHeaders, '能力項目', '能力項目名', '重み付け係数'];
    const newData = data.flatMap(row => {
      const baseData = row.slice(0, 7);
      const newRows = [];
      for (let i = 7; i < headers.length; i++) {
        const header = headers[i];
        const weightingFactor = row[i];
        if (header && (weightingFactor !== null && weightingFactor !== undefined && weightingFactor !== '')) {
          const parts = String(header).split(':');
          const id = parts[0];
          const name = parts.length > 1 ? parts.slice(1).join(':') : '';
          newRows.push([...baseData, id, name, weightingFactor]);
        }
      }
      return newRows;
    });
    setHeaders(newHeaders);
    setData(newData);
  };

  const transformAPrimeToA = () => {
    const groupedData = {};
    const newDynamicHeaders = new Set();
    data.forEach(row => {
      const key = row.slice(0, 7).join('-');
      if (!groupedData[key]) {
        groupedData[key] = { base: row.slice(0, 7), items: {} };
      }
      const header = `${row[7]}:${row[8]}`;
      newDynamicHeaders.add(header);
      groupedData[key].items[header] = row[9];
    });
    const baseHeaders = headers.slice(0, 7);
    const dynamicHeaders = Array.from(newDynamicHeaders).sort((a, b) => parseInt(a.split(':')[0], 10) - parseInt(b.split(':')[0], 10));
    const newHeaders = [...baseHeaders, ...dynamicHeaders];
    const newData = Object.values(groupedData).map(group => {
      const newRow = [...group.base];
      dynamicHeaders.forEach(header => newRow.push(group.items[header] || ''));
      return newRow;
    });
    setHeaders(newHeaders);
    setData(newData);
  };

  const transformB_namesOnly = () => {
    const newHeaders = ['時間割コード', '時間割担当教員'];
    const teacherIndex = headers.indexOf('時間割担当教員');
    const timeCodeIndex = headers.indexOf('時間割コード');
    if (teacherIndex === -1 || timeCodeIndex === -1) { alert('必要な列が見つかりません。'); return; }
    const finalData = data.flatMap(row => {
      const teachers = (row[teacherIndex] || '').split(',');
      return teachers.map(teacher => [(row[timeCodeIndex] || ''), (String(teacher).split(':')[1] || teacher).trim()]);
    });
    setHeaders(newHeaders);
    setData(finalData);
  };

  const transformB_withCode = () => {
    const newHeaders = ['時間割コード', '教員コード', '時間割担当教員'];
    const teacherIndex = headers.indexOf('時間割担当教員');
    const timeCodeIndex = headers.indexOf('時間割コード');
    if (teacherIndex === -1 || timeCodeIndex === -1) { alert('必要な列が見つかりません。'); return; }
    const finalData = data.flatMap(row => {
      const teachers = (row[teacherIndex] || '').split(',');
      return teachers.map(teacher => {
        const parts = String(teacher).trim().split(':');
        return [(row[timeCodeIndex] || ''), parts[0], parts.length > 1 ? parts.slice(1).join(':') : ''];
      });
    });
    setHeaders(newHeaders);
    setData(finalData);
  };

  return (
    <Container maxWidth="xl" sx={{ py: 4 }}>
      <Typography variant="h4" component="h1" gutterBottom>
        Excel データ変換ツール
      </Typography>
      <Card component={Paper} elevation={3}>
        <CardContent>
          <Grid container spacing={4}>
            {/* Step 1 */}
            <Grid item xs={12} md={4}>
              <Typography variant="h6" gutterBottom>ステップ1: データ読み込み</Typography>
              <Divider sx={{ mb: 2 }} />
              <Box mb={2}>
                <input accept=".xlsx, .xls" style={{ display: 'none' }} id="upload-file-button" type="file" onChange={handleFileUpload} />
                <label htmlFor="upload-file-button">
                  <Button fullWidth variant="contained" component="span" startIcon={<UploadFileIcon />}>PCからファイルを選択</Button>
                </label>
              </Box>
              <Button fullWidth variant="outlined" component="span" startIcon={<NoteAddIcon />} onClick={handleLoadReferenceFile}>参考ファイルを読み込む</Button>
              {workbook && (
                <FormControl fullWidth sx={{ mt: 2 }}>
                  <InputLabel>シートを選択</InputLabel>
                  <Select value={selectedSheet} label="シートを選択" onChange={(e) => handleSheetSelect(e.target.value, workbook)}>
                    {sheetNames.map(name => <MenuItem key={name} value={name}>{name}</MenuItem>)}
                  </Select>
                </FormControl>
              )}
            </Grid>
            {/* Step 2 */}
            <Grid item xs={12} md={4}>
              <Typography variant="h6" gutterBottom>ステップ2: 変換</Typography>
              <Divider sx={{ mb: 2 }} />
              <Box display="flex" flexDirection="column" gap={1.5}>
                <Button variant="outlined" disabled={!data.length} onClick={transformAtoAPrime}>データA → A'（展開）</Button>
                <Button variant="outlined" disabled={!data.length} onClick={transformAPrimeToA}>データA' → A（集約）</Button>
                <Button variant="outlined" disabled={!data.length} onClick={transformB_namesOnly}>データB → B'（教員名のみ）</Button>
                <Button variant="outlined" disabled={!data.length} onClick={transformB_withCode}>データB → B'（教員コード付き）</Button>
              </Box>
            </Grid>
            {/* Step 3 */}
            <Grid item xs={12} md={4}>
              <Typography variant="h6" gutterBottom>ステップ3: 書き出し</Typography>
              <Divider sx={{ mb: 2 }} />
              <Box display="flex" flexDirection="column" gap={1.5}>
                <Button variant="contained" color="primary" disabled={!data.length} onClick={() => handleExport('xlsx')} startIcon={<DownloadIcon />}>Excel形式でエクスポート</Button>
                <Button variant="contained" color="secondary" disabled={!data.length} onClick={() => handleExport('csv')} startIcon={<DownloadIcon />}>CSV形式でエクスポート</Button>
              </Box>
            </Grid>
          </Grid>
        </CardContent>
      </Card>

      {data.length > 0 && (
        <TableContainer component={Paper} sx={{ mt: 4 }} elevation={3}>
          <Table>
            <TableHead>
              <TableRow>
                {headers.map((h, index) => <TableCell key={index}><b>{h}</b></TableCell>)}
              </TableRow>
            </TableHead>
            <TableBody>
              {data.map((row, rowIndex) => (
                <StyledTableRow key={rowIndex}>
                  {row.map((cell, cellIndex) => <TableCell key={cellIndex}>{cell}</TableCell>)}
                </StyledTableRow>
              ))}
            </TableBody>
          </Table>
        </TableContainer>
      )}
    </Container>
  );
}

export default App;