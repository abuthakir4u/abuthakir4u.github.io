import React, { Component } from 'react';
import './UserOutput.css';

class UserOutput extends Component {
    render() {
        return (
            <div className="UserOutput">
                <p>This is paragraph 1</p>
                <p>{this.props.username}</p>
            </div>
        );
    }
}

export default UserOutput;