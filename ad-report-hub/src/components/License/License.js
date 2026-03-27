import React from 'react';
import './License.css';

const License = () => {
  const currentYear = new Date().getFullYear();

  return (
    <div className="license-page">
      <h1>License & Legal Information</h1>

      <div className="license-section card">
        <h2>Copyright</h2>
        <div className="section-content">
          <p className="copyright-text">
            Copyright © {currentYear} Paramount Global. All rights reserved.
          </p>
          <p>
            This software and associated documentation files (the "Software") are 
            proprietary and confidential to Paramount Global and its affiliates.
          </p>
          <p>
            Unauthorized copying, modification, distribution, or use of this Software, 
            via any medium is strictly prohibited without the express written permission 
            of Paramount Global.
          </p>
          <p>
            For licensing inquiries, please contact Paramount Global.
          </p>
        </div>
      </div>

      <div className="license-section card">
        <h2>Disclaimer</h2>
        <div className="section-content">
          <p>
            This software was developed by the Paramount Identity Architecture Team. 
            It is intended for internal use within Paramount Global and its affiliates.
          </p>
          <p>
            Any external use, distribution, or modification requires explicit written 
            authorization from Paramount Global.
          </p>
          <p>
            THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
            IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
            FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
          </p>
        </div>
      </div>

      <div className="license-section card">
        <h2>Usage Terms</h2>
        <div className="section-content">
          <ul>
            <li>This tool is for authorized personnel only</li>
            <li>All activity is logged for security and audit purposes</li>
            <li>Do not share access credentials with unauthorized users</li>
            <li>Report any security concerns to the Identity Architecture Team</li>
            <li>Data exported from this tool must be handled according to data protection policies</li>
          </ul>
        </div>
      </div>
    </div>
  );
};

export default License;
