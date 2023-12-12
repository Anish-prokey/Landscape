import h_blue_logo from "../img/h_blue_logo.png";
import h_white_logo from "../img/h_white_logo.png";
import v_blue_logo from "../img/v_blue_logo.png";
import v_white_logo from "../img/v_white_logo.png";
import profile_logo from "../img/profile.png";
import 'bootstrap/dist/css/bootstrap.css';
import 'bootstrap/dist/js/bootstrap.js';
import '../css/Home.css'
import { useRef, useEffect, useState } from "react";
import Header from "./Header"
import jsPDF from "jspdf";
import "jspdf-autotable";
import { useDownloadExcel } from "react-export-table-to-excel"
import React from 'react';
import axios from 'axios';
import LoadingBar from 'react-top-loading-bar';
import ProductModal from './ProductModal';
import '../css/ProductModal.css';
// import OData from 'react-odata';
 
import { Link, useNavigate } from 'react-router-dom';
 
export default function Home() {
  // active navigation item
  const [activeNavItem, setActiveNavItem] = useState('landscape');
 
  const handleNavItemClick = (navItem) => {
    setActiveNavItem(navItem);
  };
 
  // table search
  const [searchTerm, setSearchTerm] = useState('');
  const [data, setData] = useState([]);
  const handleSearch = (event) => {
    setSearchTerm(event.target.value);
  };
 
  const filteredData = data.filter((info) =>
    Object.values(info).some(
      (value) =>
        typeof value === 'string' &&
        value.toLowerCase().includes(searchTerm.toLowerCase())
    ));
 
  // fade alert box after 3 seconds
  const [isAlertVisible, setIsAlertVisible] = React.useState(false);
 
  const handleButtonClick = () => {
    setIsAlertVisible(true);
 
    setTimeout(() => {
      setIsAlertVisible(false);
    }, 3000);
  }
 
  const [status, setStatus] = useState(undefined);
 
  // fetching data from API and show it as table
  useEffect(() => {
    const fetchData = async () => {
      try {
        const response = await axios.get('https://sapd49.tyson.com/sap/opu/odata/sap/ZAPI_PRDVERS_SRV/LandscapeDetailsSet?sap-client=100', {
          headers: {
 
            'Content-Type': 'application/json',
 
            'Access-Control-Allow-Origin': '*',
 
            'Access-Control-Allow-Methods': 'GET, POST, PATCH, PUT, DELETE, OPTIONS',
 
            // 'Access-Control-Allow-Headers': 'Origin, Content-Type, X-Auth-Token',
          },
 
          withCredentials: true,
          auth: {
 
            username: 'JA',
            password: 'Aishh@12',
          },
        });
        // setData(response.data.feed.entry.map((item) => ({
        setData(response.data.d.results.map((item) => ({
 
          Sysid: item.Sysid,
 
          Runhdb: item.Runhdb,
 
          Sapproduct: item.Sapproduct,
 
          Systype: item.Systype,
 
        })));
      } catch (error) {
        console.error(error);
      }
    };
    fetchData();
  }, []);
  // });
 
 
 
 
  // exporting table as excel file
 
 
  const [fileFormat, setFileFormat] = useState('excel');
  const [topping, setTopping] = useState('excel');
 
  const onOptionChange = (e) => {
    setTopping(e.target.value);
    setFileFormat(e.target.value);
  };
 
  const tableref = useRef(null);
  const { onDownload } = useDownloadExcel({
    currentTableRef: tableref.current,
    filename: 'lanscape', // Specify the correct file extension (.xlsx for Excel)
    sheet: 'landscapeData', // Specify the correct sheet name
  });
 
  const ondownload = () => {
    handleButtonClick();
 
    if (topping === 'excel') {
      const formattedData = filteredData.map((info) => ({
        'SAP System ID': info.Sysid,
        'Type (Abap, dual stack, Java, Hana XS, S/4 Hana)': info.Systype,
        'Does It Run On A Hana Database': info.Runhdb,
      }));
 
      onDownload(formattedData);
      setStatus({ type: 'success' });
 
  } else if (topping === 'pdf') {
    event.preventDefault();
    setStatus({ type: 'success' });
 
    const unit = 'pt';
    const size = 'A4'; // Use A1, A2, A3, or A4
    const orientation = 'portrait'; // portrait or landscape
    const marginLeft = 20;
    const doc = new jsPDF(orientation, unit, size);
 
    doc.setFontSize(15);
    const title = '';
 
    const headers = [['SAP System ID', 'Type', 'Does It Run On A Hana Database']];
    const data = filteredData.map((user) => [user.Sysid, user.Systype, user.Runhdb]);
 
    const tableHeight = doc.autoTableEndPosY() - doc.autoTable.previous.finalY;
    const contentHeight = doc.internal.pageSize.height - 2 * marginLeft;
    const startY = (contentHeight - tableHeight) / 2;
 
    let content = {
      startY: startY > 40 ? startY : 40,
      head: headers,
      body: data,
    };
 
    doc.text(title, marginLeft, 40);
    doc.autoTable(content);
 
    doc.save('report.pdf');
  }
};
 
 
 
  const [progress, setProgress] = useState(0)
  const [recipientEmail, setRecipientEmail] = useState('');
  const [emailError, setEmailError] = useState('');
 
  // const validateEmail = (email) => {
  //   var regex = /^[a-zA-Z]+(\.[a-zA-Z]+)@tyson\.com$/;
  //   return regex.test(email);
  // };
  const validateEmail = (email) => {
    var regex = /^[a-zA-Z]+(\.[a-zA-Z]+)@tyson\.com$/;
    return regex.test(email);
  };
 
 
 
  const handleEmailChange = (event) => {
    const email = event.target.value;
    setRecipientEmail(email);
   
    if (!validateEmail(email)) {
      setEmailError('Enter Valid Tyson email id');
    } else {
      setEmailError('');
    }
  }
 
  const handleGenerateAndSend = async () => {
    event.preventDefault();
 
    if (!validateEmail(recipientEmail)) {
      setEmailError('Invalid email address. Please enter a valid tyson address');
      return;
    }
 
    try {
      setProgress(progress + 50);
     
      const response = await axios.post('http://localhost:3000/generate-and-send', {
        recipientEmail,
        fileFormat,
        filteredData
      });
 
      if (response.data.message === 'sent') {
        setStatus({ type: 'sent' });
        setProgress(100);
        handleButtonClick();
      } else {
        setProgress(100);
        setStatus({ type: 'not_sent' });
        handleButtonClick();
      }
   
    } catch (error) {
      setProgress(100);
      setStatus({ type: 'not_sent' });
      handleButtonClick();
      console.error(error);
    }
  };
 
 
  const [open, setOpen] = React.useState(false);
  const handleOpen = () => setOpen(true);
  const handleClose = () => {
    setCurrentSelectedProduct(-1);
    setOpen(false);
  };
  const [currentSelectedProduct, setCurrentSelectedProduct] = useState(null);
 
  const handleProductClick = (event, productId) => {
    event.preventDefault();
    setCurrentSelectedProduct(-1);
    handleOpen();
    setCurrentSelectedProduct(productId);
    console.log(productId)
  }
 
  return (
    <>
 
<nav className="nav navbar navbar-expand-lg navbar-light bg-light py-0">
    <div className="nav container-fluid">
        <a className="navbar-brand" href="https://www.tysonfoods.com/" target="_blank">
            <img src={h_blue_logo} class="logo img-fluid rounded-top" alt="tyson logo" />
        </a>
        <button className="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarSupportedContent" aria-controls="navbarSupportedContent" aria-expanded="false" aria-label="Toggle navigation">
            <span className="navbar-toggler-icon"></span>
        </button>
        <div className="collapse navbar-collapse" id="navbarSupportedContent">
            <ul className="navbar-nav me-auto mb-2 mb-lg-0">
                <li className="nav-item">
                    <a className={`nav-link mt-2 mx-5 ${activeNavItem === 'landscape' ? 'active' : ''}`}
                       onClick={() => handleNavItemClick('landscape')} aria-current="page" href="/landscape"
                    >SAP Landscape</a>
                </li>
                <li className="nav-item">
                    <a className={`nav-link mt-2 mx-5 ${activeNavItem === 'statistics' ? 'active' : ''}`}
                       onClick={() => handleNavItemClick('statistics')} href="../statistics/index.html" >Statistics</a>
                </li>
            </ul>
            <li className="d-flex ms-auto">
                <input className="form-control input-search me-2 mt-2 mx-5" type="search" placeholder="Search" aria-label="Search" value={searchTerm}
                       onChange={handleSearch} />
            </li>
        </div>
    </div>
</nav>
 
 
      {isAlertVisible && status?.type === 'success' && (
        <div className="alert alert-success alert-dismissible fade show" role="alert">
          <strong>"Download Successful!!!"</strong>
        </div>
      )}
      {isAlertVisible && status?.type === 'error' && (
        <div className="alert alert-warning alert-dismissible fade show" role="alert">
          <strong>Download Failed!!!</strong>
        </div>
      )}
      {isAlertVisible && status?.type === 'sent' && (
        <div className="alert alert-success alert-dismissible fade show" role="alert">
          <strong>"Email Sent Successfully!!!"</strong>
        </div>
      )}
      {isAlertVisible && status?.type === 'not_sent' && (
        <div className="alert alert-warning alert-dismissible fade show" role="alert">
          <strong>"Failed to send email!!!"</strong>
        </div>
      )}
      <LoadingBar
        color='#002554' height="3px" shadow="true"
        progress={progress}
        onLoaderFinished={() => setProgress(0)}
      />
 
      <div className="table_container my-2 mx-2">
        <div className="table-responsive ">
          <table className="table table-responsive table-bordered responsive-lg responsive-md responsive-sm"
            ref={tableref}>
            <thead>
              <tr>
                {/* <th>SAP Product</th> */}
                <th>SAP System ID </th>
                <th>Type (Abap, dual stack, Java, Hana XS, S/4 Hana)</th>
                <th>Does It Run On A Hana Database </th>
 
              </tr>
            </thead>
            <tbody>
              {filteredData.map((info) => {
                return (
                  <tr key={info.id}>
                    <td onClick={(e) => handleProductClick(e, info.Sysid)}><a ><svg width="15" height="15" fill="currentColor" className="bi bi-info-circle-fill" viewBox="0 0 16 16">
 
                      <path d="M8 16A8 8 0 1 0 8 0a8 8 0 0 0 0 16zm.93-9.412-1 4.705c-.07.34.029.533.304.533.194 0 .487-.07.686-.246l-.088.416c-.287.346-.92.598-1.465.598-.703 0-1.002-.422-.808-1.319l.738-3.468c.064-.293.006-.399-.287-.47l-.451-.081.082-.381 2.29-.287zM8 5.5a1 1 0 1 1 0-2 1 1 0 0 1 0 2z" />
 
                    </svg></a>&nbsp;{info.Sysid}</td>
 
                    {/* <td>{info.Sysid}</td> */}
 
 
 
                    <td>{info.Systype}</td>
 
 
 
                    <td>{info.Runhdb}</td>
 
 
                  </tr>
                )
              }
              )}
 
 
            </tbody>
 
 
          </table>
 
        </div>
        <hr
          style={{
            background: '#65686B',
            color: '#65686B',
            borderColor: '#65686B',
            height: '0.1rem',
          }}
        />
 
      </div>
 
 
     
        <form className="send_form mx-2 mx-4" action="" method="post">
      <div className="mb-3 mt-4">
        <input
         type="email" className="form-control small-placeholder center-input" name="" id="" aria-describedby="emailHelpId"
          placeholder="name@tyson.com"
          value={recipientEmail}
          onChange={handleEmailChange}
        />
        {emailError && <div className="error-message">{emailError}</div>}
      </div>
        <h4 className="form-text mt-4" >Send data as</h4>
 
 
        <div className="form-check" value={fileFormat}>
          <input className="form-check-input" type="radio" name="radio" id="radio_excel" value="excel" onChange={onOptionChange} checked={topping === "excel"} />
          <label className="form-check-label" htmlFor="radio_excel" >
             Excel
            </label>
        </div>
       
        <div className="form-check" value={fileFormat}>
          <input className="form-check-input" type="radio" name="radio" id="radio_pdf" value="pdf" onChange={onOptionChange} checked={topping === "pdf"}/>
          <label className="form-check-label" for="radio_pdf">
            Pdf
          </label>
        </div>
 
        <button onClick={ondownload} type="button" className="sub_btn btn mt-4 mx-2" >
          <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" className="bi bi-download" viewBox="0 0 16 16">
            <path d="M.5 9.9a.5.5 0 0 1 .5.5v2.5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-2.5a.5.5 0 0 1 1 0v2.5a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2v-2.5a.5.5 0 0 1 .5-.5z" />
            <path d="M7.646 11.854a.5.5 0 0 0 .708 0l3-3a.5.5 0 0 0-.708-.708L8.5 10.293V1.5a.5.5 0 0 0-1 0v8.793L5.354 8.146a.5.5 0 1 0-.708.708l3 3z" />
          </svg>&nbsp; Download</button>
        <button onClick={handleGenerateAndSend} type="submit" className="sub_btn btn mt-4"><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" className="bi bi-send-fill" viewBox="0 0 16 16">
          <path d="M15.964.686a.5.5 0 0 0-.65-.65L.767 5.855H.766l-.452.18a.5.5 0 0 0-.082.887l.41.26.001.002 4.995 3.178 3.178 4.995.002.002.26.41a.5.5 0 0 0 .886-.083l6-15Zm-1.833 1.89L6.637 10.07l-.215-.338a.5.5 0 0 0-.154-.154l-.338-.215 7.494-7.494 1.178-.471-.47 1.178Z" />
        </svg>&nbsp; Send</button>
 
      </form>
 
 
 
      <ProductModal handleClose={handleClose} open={open} sapSystemId={currentSelectedProduct} />
   
    </>
  );
   }