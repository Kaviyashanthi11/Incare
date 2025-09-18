import React, { useState } from "react";
import PropTypes from "prop-types";
import {
  AppBar,
  Toolbar,
  Typography,
  Drawer,
  CssBaseline,
  Box,
  List,
  ListItem,
  ListItemButton,
  ListItemText,
  IconButton,
  Divider,
} from "@mui/material";
import MenuIcon from "@mui/icons-material/Menu";
import FileUpload from "./File";
import Ecw from "./ecw";
import Incare from "./Incare";
import Avatar from "@mui/material/Avatar";

const drawerWidth = 230;

const Layout = (props) => {
  const { window } = props;
  const [mobileOpen, setMobileOpen] = useState(false);
  const [open, setOpen] = useState(true);
  const [selectedMenu, setSelectedMenu] = useState(null);

  const handleDrawerToggle = () => {
    setOpen((prev) => !prev);
    setMobileOpen((prev) => !prev);
  };

  const handleMenuClick = (menu) => {
    setSelectedMenu(menu);
    setMobileOpen(false);
  };

  const drawer = (
    <Box
      sx={{
        height: "100%",
        background: "linear-gradient(to bottom,rgb(33, 212, 243),rgb(246, 100, 132))",
        color: "white",
        display: "flex",
        flexDirection: "column",
      }}
    >
      <Toolbar sx={{ display: "flex", alignItems: "center", gap: 1 }}>
        <Avatar
          src="/images/auto1.png"
          alt="Logo"
          sx={{ width: 30, height: 30 }}
        />
        <Typography variant="h6" sx={{ fontWeight: "bold" }}>
          Menu
        </Typography>
      </Toolbar>
      <Divider sx={{ bgcolor: "rgb(33, 212, 243)" }} />

      <List>
        {[
          { name: "MSYS", icon: "/images/msys1.png", value: "fileUpload" },
          { name: "ECW Notes Update", icon: "/images/ecw2.jpg", value: "fileReading" },
          { name: "ECW Demographics", icon: "/images/Incar.png", value: "fileWriting" },
        ].map((item) => (
          <ListItem key={item.value} disablePadding>
            <ListItemButton
              onClick={() => handleMenuClick(item.value)}
              sx={{
                "&:hover": { backgroundColor: "rgba(255, 255, 255, 0.2)" },
              }}
            >
              <Avatar src={item.icon} alt={item.name} sx={{ mr: 2 }} />
              <ListItemText
              primary={
                <Typography sx={{ fontSize: "0.9rem", color: "black",fontWeight:'bold',fontFamily:'sans-serif' }}>
                  {item.name}
                </Typography>
              }
            />

            </ListItemButton>
          </ListItem>
        ))}
      </List>
    </Box>
  );

  const container = window !== undefined ? () => window().document.body : undefined;

  return (
    <Box sx={{ display: "flex" }}>
      <CssBaseline />

      {/* AppBar */}
      <AppBar
        position="fixed"
        sx={{
          zIndex: 1201,
          background: "rgb(33, 212, 243)",
          color: "white",
          boxShadow: "0px 4px 10px rgba(0,0,0,0.2)",
        }}
      >
        <Toolbar sx={{ display: "flex", alignItems: "center", gap: 2 }}>
          {/* Sidebar Toggle Button */}
          <IconButton
            color="inherit"
            aria-label="open drawer"
            edge="start"
            onClick={handleDrawerToggle}
            sx={{ display: { sm: "block" },color:'black' }}
          >
            <MenuIcon />
          </IconButton>

          {/* Logo */}
          <Avatar
            src="/images/auto1.png"
            alt="Logo"
            sx={{ width: 30, height: 30, borderRadius: "50%" }}
          />
          <Typography
            variant="h6"
            noWrap
            component="div"
            sx={{ fontWeight: "bold",color:'black' }}
          >
            Process Automation
          </Typography>
        </Toolbar>
      </AppBar>

      {/* Sidebar Drawer */}
      <Box
        component="nav"
        sx={{ width: { sm: drawerWidth }, flexShrink: { sm: 0 } }}
        aria-label="menu"
      >
        {/* Mobile Sidebar */}
        <Drawer
          container={container}
          variant="temporary"
          open={mobileOpen}
          onClose={handleDrawerToggle}
          ModalProps={{ keepMounted: true }}
          sx={{
            display: { xs: "block", sm: "none" },
            "& .MuiDrawer-paper": { boxSizing: "border-box", width: drawerWidth },
          }}
        >
          {drawer}
        </Drawer>

        {/* Desktop Sidebar */}
        <Drawer
          variant="persistent"
          sx={{
            display: { xs: "none", sm: "block" },
            "& .MuiDrawer-paper": { boxSizing: "border-box", width: drawerWidth },
          }}
          open={open}
        >
          {drawer}
        </Drawer>
      </Box>

      {/* Main Content */}
      <Box
        component="main"
        sx={{
          flexGrow: 1,
          p: 3,
          width: { sm: `calc(100% - ${open ? drawerWidth : 0}px)` },
          transition: "width 0.3s",
          mt: 8,
        }}
      >
        {selectedMenu === "fileUpload" && <FileUpload />}
        {selectedMenu === "fileReading" && <Ecw />}
        {selectedMenu === "fileWriting" && <Incare />}
      </Box>
    </Box>
  );
};

Layout.propTypes = {
  window: PropTypes.func,
};

export default Layout;


// import React, { useState } from "react";
// import PropTypes from "prop-types";
// import {
//   AppBar,
//   Toolbar,
//   Typography,
//   Drawer,
//   CssBaseline,
//   Box,
//   List,
//   ListItem,
//   ListItemButton,
//   ListItemText,
//   IconButton,
//   Divider,
// } from "@mui/material";
// import MenuIcon from "@mui/icons-material/Menu";
// import FileUpload from "./File"; // File Upload Component
// import CloudUploadIcon from "@mui/icons-material/CloudUpload";
// import MenuBookIcon from "@mui/icons-material/MenuBook";
// import Ecw from "./ecw";
// import Incare from "./Incare";
// import EditIcon from "@mui/icons-material/Edit"

// const drawerWidth = 200;

// const Layout = (props) => {
//   const { window } = props;
//   const [mobileOpen, setMobileOpen] = useState(false);
//   const [open, setOpen] = useState(true); // Sidebar open by default on large screens
//   const [selectedMenu, setSelectedMenu] = useState(null); // Default no component displayed

//   // Toggle Sidebar for all screen sizes
//   const handleDrawerToggle = () => {
//     setOpen((prev) => !prev);
//     setMobileOpen((prev) => !prev);
//   };

//   // Handle menu selection
//   const handleMenuClick = (menu) => {
//     setSelectedMenu(menu);
//     setMobileOpen(false); // Close drawer on mobile after selection
//   };

//   // Sidebar Drawer
//   const drawer = (
//     <div>
//       <Toolbar />
//       <Divider />
//       <List  sx={{ paddingTop: 0 }}>
//         <ListItem disablePadding>
//           <ListItemButton onClick={() => handleMenuClick("fileUpload")}>
//             <CloudUploadIcon sx={{ mr: 2 }} />
//             <ListItemText primary="MSYS" />
//           </ListItemButton>
//         </ListItem>
//         <ListItem disablePadding>
//           <ListItemButton onClick={() => handleMenuClick("fileReading")}>
//             <MenuBookIcon sx={{ mr: 2 }} />
//             <ListItemText primary="ECW NOTES UPDATE" />
//           </ListItemButton>
//         </ListItem>
//         <ListItem disablePadding>
//           <ListItemButton onClick={() => handleMenuClick("fileWriting")}>
//             <EditIcon sx={{ mr: 2 }} />
//             <ListItemText primary="ECW DEMOGRAPHICS" />
//           </ListItemButton>
//         </ListItem>
//       </List>
//     </div>
//   );

//   const container = window !== undefined ? () => window().document.body : undefined;

//   return (
//     <Box sx={{ display: "flex" }}>
//       <CssBaseline />

//       {/* App Bar */}
//       <AppBar position="fixed" sx={{ zIndex: 1201, backgroundColor: "skyblue",color:'black', }}>
//         <Toolbar>
//           {/* Hamburger Menu Icon for Sidebar Toggle */}
//           <IconButton
//             color="inherit"
//             aria-label="open drawer"
//             edge="start"
//             onClick={handleDrawerToggle}
//             sx={{ mr: 2, display: { sm: "block" } }} // Show on all screens
//           >
//             <MenuIcon />
//           </IconButton>

//           <Typography variant="h6" noWrap component="div" style={{ color: "black"}}>
//             Process Automation
//           </Typography>
//         </Toolbar>
//       </AppBar>

//       {/* Sidebar */}
//       <Box
//         component="nav"
//         sx={{ width: { sm: drawerWidth }, flexShrink: { sm: 0 } }}
//         aria-label="menu"
//       >
//         {/* Mobile Drawer */}
//         <Drawer
//           container={container}
//           variant="temporary"
//           open={mobileOpen}
//           onClose={handleDrawerToggle}
//           ModalProps={{ keepMounted: true }}
//           sx={{
//             display: { xs: "block", sm: "none" },
//             "& .MuiDrawer-paper": { boxSizing: "border-box", width: drawerWidth },
//           }}
//         >
//           {drawer}
//         </Drawer>

//         {/* Permanent Sidebar for Larger Screens */}
//         <Drawer
//           variant="persistent"
//           sx={{
//             display: { xs: "none", sm: "block" },
//             "& .MuiDrawer-paper": { boxSizing: "border-box", width: drawerWidth },
//           }}
//           open={open} // Sidebar toggles on large screens
//         >
//           {drawer}
//         </Drawer>
//       </Box>

//       {/* Main Content */}
//       <Box
//         component="main"
//         sx={{
//           flexGrow: 1,
//           p: 3,
//           width: { sm: `calc(100% - ${open ? drawerWidth : 0}px)` },
//           transition: "width 0.3s",
//           mt: 8,
//         }}
//       >
//         {selectedMenu === "fileUpload" && <FileUpload />}
//         {selectedMenu === "fileReading" && <Ecw />}
//         {selectedMenu === "fileWriting" && <Incare />}
//       </Box>
//     </Box>
//   );
// };

// Layout.propTypes = {
//   window: PropTypes.func,
// };

// export default Layout;





// import React, { useState } from "react";
// import PropTypes from "prop-types";
// import {
//   AppBar,
//   Toolbar,
//   Typography,
//   Drawer,
//   CssBaseline,
//   Box,
//   List,
//   ListItem,
//   ListItemButton,
//   ListItemText,
//   IconButton,
//   Divider,
// } from "@mui/material";
// import MenuIcon from "@mui/icons-material/Menu";
// import FileUpload from "./File";
// import Ecw from "./ecw";
// import Incare from "./Incare";
// import Avatar from "@mui/material/Avatar";

// const drawerWidth = 230;

// const Layout = (props) => {
//   const { window } = props;
//   const [mobileOpen, setMobileOpen] = useState(false);
//   const [open, setOpen] = useState(true);
//   const [selectedMenu, setSelectedMenu] = useState(null);

//   const handleDrawerToggle = () => {
//     setOpen((prev) => !prev);
//     setMobileOpen((prev) => !prev);
//   };

//   const handleMenuClick = (menu) => {
//     setSelectedMenu(menu);
//     setMobileOpen(false);
//   };

//   const drawer = (
//     <Box
//       sx={{
//         height: "100%",
//         background: "linear-gradient(to bottom,rgb(33, 212, 243),rgb(246, 100, 132))",
//         color: "white",
//         display: "flex",
//         flexDirection: "column",
//       }}
//     >
//       <Toolbar sx={{ display: "flex", alignItems: "center", gap: 1 }}>
//         <Avatar
//           src="/images/auto1.png"
//           alt="Logo"
//           sx={{ width: 40, height: 40 }}
//         />
//         <Typography variant="h6" sx={{ fontWeight: "bold" }}>
          
//         </Typography>
//       </Toolbar>
//       <Divider sx={{ bgcolor: "rgb(33, 212, 243)" }} />

//       <List>
//         {[
//           { name: "MSYS", icon: "/images/msys2.png", value: "fileUpload" },
//           { name: "ECW Notes Update", icon: "/images/ecw1.jpg", value: "fileReading" },
//           { name: "ECW Demographics", icon: "/images/ecw.jpg", value: "fileWriting" },
//         ].map((item) => (
//           <ListItem key={item.value} disablePadding>
//             <ListItemButton
//               onClick={() => handleMenuClick(item.value)}
//               sx={{
//                 "&:hover": { backgroundColor: "rgba(255, 255, 255, 0.2)" },
//               }}
//             >
//               <Avatar src={item.icon} alt={item.name} sx={{ mr: 2 }} />
//               <ListItemText
//               primary={
//                 <Typography sx={{ fontSize: "0.9rem", color: "black",fontWeight:'bold',fontFamily:'sans-serif' }}>
//                   {item.name}
//                 </Typography>
//               }
//             />

//             </ListItemButton>
//           </ListItem>
//         ))}
//       </List>
//     </Box>
//   );

//   const container = window !== undefined ? () => window().document.body : undefined;

//   return (
//     <Box sx={{ display: "flex" }}>
//       <CssBaseline />

//       {/* AppBar */}
//       <AppBar
//         position="fixed"
//         sx={{
//           zIndex: 1201,
//           background: "rgb(33, 212, 243)",
//           color: "white",
//           boxShadow: "0px 4px 10px rgba(0,0,0,0.2)",
//         }}
//       >
//         <Toolbar sx={{ display: "flex", alignItems: "center", gap: 2 }}>
//           {/* Sidebar Toggle Button */}
//           <IconButton
//             color="inherit"
//             aria-label="open drawer"
//             edge="start"
//             onClick={handleDrawerToggle}
//             sx={{ display: { sm: "block" },color:'black' }}
//           >
//             <MenuIcon />
//           </IconButton>

//           {/* Logo */}
//           <Avatar
//             src="/images/auto1.png"
//             alt="Logo"
//             sx={{ width: 30, height: 30, borderRadius: "50%" }}
//           />
//           <Typography
//             variant="h6"
//             noWrap
//             component="div"
//             sx={{ fontWeight: "bold",color:'black' }}
//           >
//             Process Automation
//           </Typography>
//         </Toolbar>
//       </AppBar>

//       {/* Sidebar Drawer */}
//       <Box
//         component="nav"
//         sx={{ width: { sm: drawerWidth }, flexShrink: { sm: 0 } }}
//         aria-label="menu"
//       >
//         {/* Mobile Sidebar */}
//         <Drawer
//           container={container}
//           variant="temporary"
//           open={mobileOpen}
//           onClose={handleDrawerToggle}
//           ModalProps={{ keepMounted: true }}
//           sx={{
//             display: { xs: "block", sm: "none" },
//             "& .MuiDrawer-paper": { boxSizing: "border-box", width: drawerWidth },
//           }}
//         >
//           {drawer}
//         </Drawer>

//         {/* Desktop Sidebar */}
//         <Drawer
//           variant="persistent"
//           sx={{
//             display: { xs: "none", sm: "block" },
//             "& .MuiDrawer-paper": { boxSizing: "border-box", width: drawerWidth },
//           }}
//           open={open}
//         >
//           {drawer}
//         </Drawer>
//       </Box>

//       {/* Main Content */}
//       <Box
//   component="main"
//   sx={{
//     flexGrow: 1,
//     p: 3,
//     width: { 
//       xs: '100%',
//       sm: `calc(100% - ${open ? drawerWidth : 0}px)` 
//     },
//     marginLeft: { 
//       xs: 0,
//       sm: open ? `${drawerWidth}px` : 0 
//     },
//     transition: theme => theme.transitions.create(['width', 'margin'], {
//       easing: theme.transitions.easing.sharp,
//       duration: theme.transitions.duration.enteringScreen,
//     }),
//     mt: 8,
//     display: "flex",
//     justifyContent: "left",
//     alignItems: "left",
//     position: "relative",
//     height: "calc(100vh - 64px)",
//     "&::before": {
//       content: '""',
//       position: "absolute",
//       top: 0,
//       left: 0,
//       right: 0,
//       bottom: 0,
//       backgroundImage: "url('/images/sample.png')",
//       backgroundSize: "cover",
//       backgroundPosition: "center",
//       backgroundRepeat: "no-repeat",
//       zIndex: -1,
//       marginLeft:'-235px',
//       marginTop:'-10px'
//     }
//   }}
// >
//   {selectedMenu === "fileUpload" && <FileUpload />}
//   {selectedMenu === "fileReading" && <Ecw />}
//   {selectedMenu === "fileWriting" && <Incare />}
// </Box>
//     </Box>
//   );
// };


// Layout.propTypes = {
//   window: PropTypes.func,
// };

// export default Layout;


