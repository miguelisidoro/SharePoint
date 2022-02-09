import React from 'react';
import { Content } from './layout/Content';
import Footer from './layout/Footer';
import Header from './layout/Header';

class App extends React.Component {
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

export default App
