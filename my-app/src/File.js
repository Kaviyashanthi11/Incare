import React, { useState, useRef } from "react";
import { Button, Backdrop, CircularProgress } from "@mui/material";
import { useMediaQuery } from "@mui/material";
import { useTheme } from "@mui/material/styles";

const FileUpload = () => {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const fileInputRef = useRef(null);


  const theme = useTheme();
  const isLargeScreen = useMediaQuery(theme.breakpoints.up("md")); // md = 960px by default
  

  const handleFileChange = (event) => {
    setFile(event.target.files[0]);
  };
  const handleUpload = async () => {
    if (!file) {
        alert("Please select a file first.");
        return;
    }

    setLoading(true);

    const formData = new FormData();
    formData.append("file", file);
    formData.append("originalPath", file.name);

    try {
        // First, process the file
        const processResponse = await fetch("http://localhost:5000/process", {
            method: "POST",
            body: formData,
        });

        const result = await processResponse.json();

        if (result.success) {
            const updatedRows = result.results?.successful?.length || 0;

            if (updatedRows > 0) {
                alert(`${updatedRows} note${updatedRows > 1 ? 's' : ''} ${updatedRows > 1 ? 'were' : 'was'} successfully posted.`);
                
                // ✅ Handle file download if available
                if (result.fileId) {
                    const downloadUrl = `http://localhost:5000/download/${result.fileId}`;
                    const link = document.createElement('a');
                    link.href = downloadUrl;
                    link.setAttribute('download', 'updated_file.xlsx');
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                }
            } else {
                alert("All rows were already updated.");
            }
        } else {
            // ✅ Handle "Notes not found in Excel." message
            if (result.message === "Notes not found in Excel.") {
                alert("Notes not found in Excel.");
            } else {
                alert(`Error: ${result.error || "File processing failed."}`);
            }
        }

        // Reset file input
        setFile(null);
        if (fileInputRef.current) {
            fileInputRef.current.value = "";
        }
    } catch (error) {
        console.error("Upload Error:", error);
        alert("Backend Does not Connected.");
    } finally {
        setLoading(false);
    }
};


  return (
    <div
    style={{
      padding: "30px",
      maxWidth: "400px",
      margin: "0 auto",
      textAlign: "center",
      marginLeft: isLargeScreen ? "300px" : "auto", // Apply margin-left only for larger screens
    }}
  >
          <h2>Upload Your Excel File</h2>
      <input type="file" accept=".xlsx" onChange={handleFileChange} ref={fileInputRef} /><br/><br/>
      <Button
        variant="contained"
        onClick={handleUpload}
        style={{
          marginTop: "10px",
          padding: "10px 20px",
          backgroundColor: "pink",
          color: "black",
        }}
      >
        Upload
      </Button>
      <Backdrop open={loading} style={{ zIndex: 9999, color: "#fff" }}>
        <CircularProgress color="inherit" />
      </Backdrop>
    </div>
  );
};

export default FileUpload;

