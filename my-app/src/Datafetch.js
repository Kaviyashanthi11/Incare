import React, { useState } from 'react';
import { TextField, Button, Grid, Typography ,Link} from '@mui/material';
import companylogo from './images/companylogo.jpg'

const ARPatientSearch = () => {
  const [accountNumber, setAccountNumber] = useState('');
  const [patientData, setPatientData] = useState(null);
  const [error, setError] = useState(null);

  const handleFetchData = async () => {
    // Validation: Check if account number is exactly 6 digits
    if (!accountNumber) {
      alert('Please enter an account number.');
      return;
    }
    if (!/^\d{6}$/.test(accountNumber)) {
      alert('The Account Number Was Not Found.');
      return;
    }

    try {
      setError(null); // Reset error state before making the request
      setPatientData(null); // Reset patient data before fetching

      const response = await fetch('http://localhost:5000/fetchPatientData', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ accountNumber }),
      });

      const data = await response.json();

      if (response.ok) {
        if (data.patientName) {
          setPatientData(data);
        } else {
          setError('No data found for this account number.');
        }
      } else {
        setError(data.error || 'Failed to fetch patient data.');
      }
    } catch (err) {
      console.error('Error fetching data:', err);
      setError('Failed to fetch patient data. Please check the server connection and try again.');
    }
  };

  return (
    <Grid 
    container 
    direction="column" 
    sx={{ 
      backgroundImage: "url('/blue.jpg')", 
      backgroundSize: "cover",
      backgroundPosition: "center",
      minHeight: "100vh", 
      padding: 2 
    }}
  >
     {/* Top Section with Image & Links */}
     <Typography className="logo" sx={{ marginTop: "10px" ,color:'#1a237e',fontSize:'18px'}}>
        <span className="version" style={{ marginLeft: "5px" }}><b>System MHS Version 7.000.12</b></span>
      </Typography>
    <Grid container justifyContent="space-between" alignItems="center" sx={{ mb: 2 }}>
      {/* Left: Logo */}
      <Grid item>
        <img src={companylogo} alt="Hospital Logo" style={{ height: '50px',width:'260px' }} />
      </Grid>
      
  <Typography variant="h5" color="primary" sx={{ marginTop:'-60px',color:'#1a237e' }}>
          A/R Patient Search
        </Typography>
      {/* Right: Links */}
      <Grid item>
        <Link component="span" sx={{ mx: 1, color: 'brown', cursor: 'pointer', }}>Help</Link>
        <Link component="span" sx={{ mx: 1, color: 'brown', cursor: 'pointer' }}>Back to Menu</Link>
        <Button variant="contained"  sx={{ ml: 1,backgroundColor:'skyblue',color:'blueviolet' }}>Log Off</Button>
      </Grid>
    </Grid>

      <Grid item xs={12} md={8} sx={{ textAlign: 'left' }}>
      <Typography variant="h6" sx={{mb:7, marginTop:'-50px' ,textAlign:'center',color:'#304ffe'}}>
          DEERBROOK EMERGENCY HOSPITAL
        </Typography>
        
    <Grid container direction="column" sx={{ minHeight: '100vh', padding: 2 }}>
      <Typography variant="h5" color="primary">A/R Patient Search</Typography>
      
      {/* Account Number Input */}
      <Grid container spacing={1} alignItems="center">
        <Grid item xs={12} sm={2}>
          <Typography>Account Number:</Typography>
        </Grid>
        <Grid item xs={12} sm={3}>
          <TextField
            fullWidth
            variant="outlined"
            value={accountNumber}
            onChange={(e) => setAccountNumber(e.target.value)}
          />
        </Grid>
        <Grid item xs={12} sm={2}>
           <Button variant="contained" onClick={handleFetchData} sx={{ 
                    backgroundColor: '#bdbdbd', 
                    color: 'black', 
                    border: '1px solid black',  // Set border color
                }}>Find</Button>
        </Grid>
      </Grid>

      {/* Display error if any */}
      {error && (
        <Grid container sx={{ mt: 2 }}>
          <Grid item xs={12}>
            <Typography color="error">{error}</Typography>
          </Grid>
        </Grid>
      )}

      {/* Display patient data */}
      {patientData && (
        <Grid container spacing={2} sx={{ mt: 4 }}>
          <Grid item xs={12}>
            <Typography variant="h6">Patient Information:</Typography>
            <Typography>Name: {patientData.patientName}</Typography>
            <Typography>Gender: {patientData.gender}</Typography>
          </Grid>
        </Grid>
      )}
    </Grid>
    </Grid>
    </Grid>
  );
};

export default ARPatientSearch;

