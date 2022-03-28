import { mount, shallow } from 'enzyme';
import React from 'react';
import CipAggregator from './CipAggregator';

test('Hello', () => {
    const component = shallow(
        <CipAggregator 
            description='Hello world'
            environmentMessage='environment'
            isDarkTheme
            userDisplayName='Andrei'
            hasTeamsContext={false}
        />
    );

    const header = <h2>Well done, Andrei!</h2>;
    expect(component.html()).toMatchSnapshot()
    expect(component.contains(header)).toBeTruthy();
    expect(component.find("[data-testid='description']").text())
        .toBe("Web part property value: Hello world")
});