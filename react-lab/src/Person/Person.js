import React from 'react';
import './Person.css'; //webpack will include this in html

const person = (props) => {
    return (
        <div className="Person">
            <p onClick={props.click}>I am person... {props.name} ... {props.age}</p>
            {props.children}
            <input type="text" onChange={props.changed} value={props.name}/>
        </div>
    );
};

export default person;