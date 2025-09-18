import React from "react";
import Layout from '../src/Layout';
import { ThemeProvider, createTheme } from '@mui/material/styles';

const App = (props) => {
  const theme = createTheme(); // Create a default theme or customize it as needed
  
  return (
    <ThemeProvider theme={theme}>
      <Layout />
    </ThemeProvider>
  );
};

export default App;