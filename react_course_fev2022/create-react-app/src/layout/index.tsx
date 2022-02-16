import React from 'react';
import Content from './Content';
import Footer from './Footer';
import Header from './Header';

class AppLayout extends React.Component {
  render(){
    return (
      <>
        <Header />
        <Content />
        <Footer />
      </>
    )
  }
}

export default AppLayout
