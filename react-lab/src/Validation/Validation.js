import { Component } from "react";

class Validation extends Component {

    render() {

        let validationDisplay;
        if (this.props.strLength < 5 ){
            validationDisplay = <div>{this.props.strLength} - Text too short</div>;
        } else {
            validationDisplay = <div>{this.props.strLength} - Text long enough</div>;
        }

        return (validationDisplay);
    }
}

export default Validation;