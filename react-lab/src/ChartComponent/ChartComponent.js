import { Component } from "react";

class ChartComponent extends Component {
    render() {
        let style = {
            display: 'inline-block',
            padding: '16px',
            textAlign: 'center',
            margin: '16px',
            border: '1px solid black'
        }
        return (
            <div style={style} onClick={this.props.removeMe}>
                {this.props.letter}
            </div>
        )
    }
}

export default ChartComponent;