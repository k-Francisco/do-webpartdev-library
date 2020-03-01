import React, { createContext, useReducer, useContext, useEffect } from "react";

export const DispatchContext = createContext(null);
export const StateContext = createContext(null);

export const Store = props => {
  const [state, dispatch] = useReducer(props.reducer, props.initialState);

  return (
    <StateContext.Provider value={state}>
      <DispatchContext.Provider value={dispatch}>
        {props.children.map(child => {
          return child;
        })}
      </DispatchContext.Provider>
    </StateContext.Provider>
  );
};
