import React from 'react';
import Component1 from './components/component1';
import Component2 from './components/component2';
class App extends React.Component {
  render(){
    return (
      <>
        <Component1 />
        <Component1 />
        <Component2 />
      </>
    )
  }
}

export default App
