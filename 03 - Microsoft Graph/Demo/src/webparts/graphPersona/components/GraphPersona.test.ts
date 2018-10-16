import * as React from 'react';
import { configure, mount, ReactWrapper } from 'enzyme';
import * as Adapter from 'enzyme-adapter-react-15';
configure({ adapter: new Adapter() });

import GraphPersona from "./GraphPersona";
import { IGraphPersonaProps } from './IGraphPersonaProps';
import { IGraphPersonaState } from './IGraphPersonaState';
import { IGraphPersonaService } from './IGraphPersonaService';
describe("GraphPersona", () => {
    let reactComponent: ReactWrapper<IGraphPersonaProps, IGraphPersonaState>;
    let getProfileInfoPromiseResolve;
    let getProfileInfoPromiseReject;
    beforeEach(() => {
        // Implement test service
        let service: IGraphPersonaService = {
            getProfileInfo() {
                return new Promise((resolve, reject) => {
                    getProfileInfoPromiseResolve = resolve;
                    getProfileInfoPromiseReject = reject;
                });
            },
            getPhoto() {
                return Promise.resolve("http://myhost/myimage.jpg");
            }
        };
        // Mount react component
        reactComponent = mount(React.createElement(
            GraphPersona,
            {
                service: service
            }
        ));
    });

    afterEach(() => {
        reactComponent.unmount();
    });

    it("Should load with empty persona", () => {
        const element = reactComponent.find('.ms-Persona-primaryText');
        expect(element.length).toBeGreaterThan(0);
        expect(element.text()).toEqual("");
    });

    it("Should show real name when resolved", async () => {
      const element = reactComponent.find('.ms-Persona-primaryText');
      expect(element.length).toBeGreaterThan(0);
      expect(element.text()).toEqual("");
      getProfileInfoPromiseResolve({
        "businessPhones": [
          "+1 412 555 0109"
        ],
        "displayName": "Megan Bowen",
        "mail": "MeganB@M365x214355.onmicrosoft.com",
      });
      await reactComponent.update();
      expect(element.text()).toEqual("Megan Bowen");
    });
    it("Should possible to beautify name", async () => {
      const element = reactComponent.find('.ms-Persona-primaryText');
      expect(element.length).toBeGreaterThan(0);
      expect(element.text()).toEqual("");
      getProfileInfoPromiseResolve({
        "businessPhones": [
          "+1 412 555 0109"
        ],
        "displayName": "Megan Bowen",
        "mail": "MeganB@M365x214355.onmicrosoft.com",
      });
      await reactComponent.update();
      expect(element.text()).toEqual("Megan Bowen");
      const button = reactComponent.find("button");
      button.simulate('click');
      expect(element.text()).toEqual("Per Jakobsen");  
    });
});