import React, { Component, useState } from 'react';
import './App.css';
import Person from './Person/Person';
import UserInput from './UserInput/UserInput';
import UserOutput from './UserOutput/UserOutput';
import Validation from './Validation/Validation';
import ChartComponent from './ChartComponent/ChartComponent';

class App extends Component {

    state = {
        persons: [{ id: 'abc1', name: 'abu', age: 28 }, { id: 'abc2', name: 'kamal', age: 30 }, { id: 'abc3', name: 'anwar', age: 32 }],
        userName: 'Abuthakir',
        showPerson: false,
        strLength: 0,
        inputStringLetters: []
    };
    deletePersonHandler = (personIndex) => {
        const persons = [...this.state.persons];
        persons.splice(personIndex, 1);
        this.setState({ persons: persons });
    }
    switchNameHandler = (newName) => {
        console.log('called switchNameHandler');
        this.setState({ persons: [{ name: newName, age: 30 }, { name: 'kamal', age: 31 }, { name: 'anwar', age: 32000 }] });
    }
    nameChangeHandler = (event, id) => {
        const personIndex = this.state.persons.findIndex((person) => {
            return (person.id === id);
        })

        //Modern way of copying/cloning the object
        const person = { ...this.state.persons[personIndex] };

        //Other legacy way of copying the object
        //const person = Object.assign({}, this.state.persons[personIndex]);

        person.name = event.target.value;
        const persons = [...this.state.persons];
        persons[personIndex] = person;

        this.setState({
            persons: persons
        });
    }
    updateUsernameHandler = (event) => {
        this.setState({ userName: event.target.value })
    }
    togglePersonHandler = () => {

        const doesShow = this.state.showPerson;
        this.setState({ showPerson: !doesShow });
    }
    setInputStringLengthHandler = (event) => {
        const strLength = event.target.value.length;
        const inputStringLetters = event.target.value.split('');
        this.setState({ strLength: strLength, inputStringLetters: inputStringLetters });
    }
    removeMeHandler = (event, index) => {
        console.log('index', index);
        const inputStringLetters = [...this.state.inputStringLetters];
        inputStringLetters.splice(index, 1);
        this.setState({ inputStringLetters: inputStringLetters });
    }
    render() {
        const style = {
            backgroundColor: "blue",
            color: "#fff",
            padding: "10px"
        };
        let person = null;
        if (this.state.showPerson) {
            person = (
                <div>
                    {this.state.persons.map((person, index) => {
                        return <Person name={person.name} age={person.age} click={() => this.deletePersonHandler(index)}
                            key={person.id}
                            changed={(event) => this.nameChangeHandler(event, person.id)} />
                    })}
                </div>
            );
        }
        return (
            <div className="App" >
                <h1>Hi, i am ractjs app</h1>
                <button onClick={() => this.switchNameHandler('Abuthakir updated')} style={style}>Switch Name</button>
                <button onClick={this.togglePersonHandler} >Toggle</button>
                {person}
                <UserOutput username={this.state.userName} />
                <UserInput updateUsername={this.updateUsernameHandler} username={this.state.userName} />

                <div>
                    <input type="text" onChange={this.setInputStringLengthHandler} value={this.state.inputStringLetters.join('')} />
                    <p>Lenght of the input string: {this.state.strLength}</p>
                    <Validation strLength={this.state.strLength}></Validation>

                    {this.state.inputStringLetters.map((letter, index) => {
                        return <ChartComponent letter={letter} removeMe={(event) => this.removeMeHandler(event, index)} key={index}></ChartComponent>
                    })}

                </div>
            </div >
        );
    }
}

export default App;
