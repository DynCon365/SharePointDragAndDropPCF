import * as React from "react";
import { ServiceProvider } from "../components/pcf-react/ServiceProvider";
export const ServiceProviderContext = React.createContext<ServiceProvider>(new ServiceProvider());
