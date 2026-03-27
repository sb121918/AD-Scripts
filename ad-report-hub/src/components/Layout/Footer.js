import React from 'react';
import { Link } from 'react-router-dom';
import './Footer.css';

const Footer = () => {
  const currentYear = new Date().getFullYear();

  return (
    <footer className="footer">
      <div className="footer-content">
        <span className="copyright">
          © {currentYear} Paramount Global. All rights reserved.
        </span>
        <span className="footer-divider">|</span>
        <Link to="/license" className="footer-link">
          See License tab for disclaimer
        </Link>
      </div>
    </footer>
  );
};

export default Footer;
