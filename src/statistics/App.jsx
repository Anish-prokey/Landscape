import React, { useState, useEffect } from 'react';
import h_blue_logo from '../img/h_blue_logo.png';
import 'bootstrap/dist/css/bootstrap.css';
import 'bootstrap/dist/js/bootstrap.js';
import '../css/Home.css';
import axios from 'axios';


export default function App() {
  const [activeNavItem, setActiveNavItem] = useState('statistics');
  const [selectedSystem, setSelectedSystem] = useState(null);
  const [systems, setSystems] = useState([]);
  const[data,setData]= useState([]);

  const groupedSystems = {};
  useEffect(() => {
    const fetchData = async () => {
      try {

        const response3 = await axios.get('https://sapd49.tyson.com/sap/opu/odata/sap/ZAPI_PRDVERS_SRV/SysDetailsSet?sap-client=100', {
        headers: {

          'Content-Type': 'application/json',

          'Access-Control-Allow-Origin': '*',

          'Access-Control-Allow-Methods': 'GET, POST, PATCH, PUT, DELETE, OPTIONS',
        },

        withCredentials: true,
        auth: {
          username: 'JA',
          password: 'Aishh@12',
        },
      });
        
          response3.data.d.results.forEach((item) => {
            const key = item.Name;
          
            if (!groupedSystems[key]) {
              groupedSystems[key] = {};
            }
          
            const sysid = item.Sysid;
          
            if (!groupedSystems[key][sysid]) {
              groupedSystems[key][sysid] = { ...item };
            } else {
              const existingItem = groupedSystems[key][sysid];
              // Check if Inststatus is different, and if so, append the new Inststatus to Inststatus2
              if (item.Inststatus !== existingItem.Inststatus) {
                existingItem.Inststatus2 = existingItem.Inststatus2
                  ? `${existingItem.Inststatus2}, ${item.Inststatus}`
                  : item.Inststatus;
              }
              // Check if Version is different, and if so, append the new Version to Version2
              if (item.Version !== existingItem.Version) {
                existingItem.Version2 = existingItem.Version2
                  ? `${existingItem.Version2}, ${item.Version}`
                  : item.Version;
              }
              
            }
          });
          
          // Transform the grouped data into an array
          const groupedSystemsArray = Object.entries(groupedSystems).map(([Name, items]) => {
            const mergedItems = Object.values(items);
            return {
              Name,
               mergedItems,
            };
          });
          
          setSystems(groupedSystems);
          
          } catch (error) {
        console.error('Error fetching systems:', error);
      }
    };

    fetchData();
  }, []);

  const handleNavItemClick = (navItem) => {
    setActiveNavItem(navItem);
  };


  
  const HorizontalTimeline = ({ timelines }) => {
    useEffect(() => {
      
    }, [timelines]);
  
    return (
      <div className="timeline-container">
        {timelines &&
          Object.entries(timelines).map(([sysid, timeline], index) => (
            <div className="timeline horizontal reverse" key={index}>
              <div className="timeline-box">
                <div className="block">
                  <div className="square up">
                    <h3>{sysid}</h3>
                    <h4></h4>
                    <p>
                    <b>Version: {timeline.Inststatus === '-' ? timeline.Version : timeline.Inststatus2 === '-' ? timeline.Version2 : 'NA'}
</b>

                    </p>
                  </div>
                  <div className="circle">
                    {timeline.Inststatus === '-' || timeline.Inststatus2 === '-' ? (
                      <p>{timeline.ModDate.substring(0, 4)}</p>
                    ) : (
                      <p>NA</p>
                    )}
                  </div>
                  <div className="link"></div>
                  <div className="square down blank"></div>
                </div>
                <div className="block">
                  <div className="square up blank"></div>
                  <div className="circle activated">
                    {timeline.Inststatus === '+' || timeline.Inststatus2 === '+' ? (
                      <p>{timeline.ModDate.substring(0, 4)}</p>
                    ) : (
                      <p>NA</p>
                    )}
                  </div>
                  <div className="link"></div>
                  <div className="square down">
                    <h3>{sysid}</h3>
                    <p>
                    <b>Version: {timeline.Inststatus === '+' ? timeline.Version : timeline.Inststatus2 === '+' ? timeline.Version2 : 'NA'}
</b>

                    </p>
                  </div>
                </div>
                <div id="end"></div>
              </div>
            </div>
          ))}
      </div>
    );
  };
  
  
  const handleSystemChange = (event) => {
    setSelectedSystem(event.target.value);
  };

  return (
    <>
      <nav className="nav navbar navbar-expand-lg navbar-light bg-light py-0">
        <div className="nav container-fluid">
          <a
            className="navbar-brand"
            href="https://www.tysonfoods.com/"
            target="_blank"
          >
            <img
              src={h_blue_logo}
              className="logo img-fluid rounded-top"
              alt="tyson logo"
            />
          </a>
          <button
            className="navbar-toggler"
            type="button"
            data-bs-toggle="collapse"
            data-bs-target="#navbarSupportedContent"
            aria-controls="navbarSupportedContent"
            aria-expanded="false"
            aria-label="Toggle navigation"
          >
            <span className="navbar-toggler-icon"></span>
          </button>
          <div
            className="collapse navbar-collapse"
            id="navbarSupportedContent"
          >
            <ul className="navbar-nav me-auto mb-2 mb-lg-0">
              <li className="nav-item">
                <a
                  className={`nav-link mt-2 mx-5 ${
                    activeNavItem === 'landscape' ? 'active' : ''
                  }`}
                  onClick={() => handleNavItemClick('landscape')}
                  aria-current="page"
                  href="/landscape"
                >
                  SAP Landscape
                </a>
              </li>
              <li className="nav-item">
                <a
                  className={`nav-link mt-2 mx-5 ${
                    activeNavItem === 'statistics' ? 'active' : ''
                  }`}
                  onClick={() => handleNavItemClick('statistics')}
                  href="../statistics/index.html"
                >
                  Statistics
                </a>
              </li>
              <li className="nav-item">
                
              </li>
            </ul>
          </div>
        </div>
      </nav>
      <div className="d-flex justify-content-center">
        <div className="dropdown-container">
          <select value={selectedSystem} onChange={handleSystemChange}>
            <option value={null}>Select A Product</option>
            {Object.keys(systems).map((systemId, index) => (
              <option key={index} value={systemId}>
                {systemId}
              </option>
            ))}
          </select>
        </div>
      </div>
      {selectedSystem && (
        <div>
          <HorizontalTimeline timelines={systems[selectedSystem]} />
        </div>
      )}
    </>
  );
}