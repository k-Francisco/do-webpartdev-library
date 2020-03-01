import { sp, IItemAddResult } from "@pnp/sp/presets/all";
import { IItemOptions } from "../interfaces/IItemOptions";
import { log } from "../utils/Utilities";
import { useState, useEffect, useContext } from "react";
import { DispatchContext, StateContext } from "../components/Store";

export const useListItem = (context: any) => {
  const [items, setItems] = useState([]);
  const [isLoading, setIsLoading] = useState(false);

  const assignDefaults = (options): IItemOptions => {
    if (!options) {
      const defaultOptions: IItemOptions = {
        isId: false,
        select: "",
        filter: "",
        expand: "",
        top: 100,
        skip: 0,
        isSkipReverse: false,
        isAscending: true,
        orderBy: "Id",
        disableAutoRefresh: false
      };
      return defaultOptions;
    }

    options.select = options.select ? options.select : "";
    options.filter = options.filter ? options.filter : "";
    options.expand = options.expand ? options.expand : "";
    options.top = options.top ? options.top : 100;
    options.skip = options.skip ? options.skip : 0;
    options.isSkipReverse = options.isSkipReverse
      ? options.isSkipReverse
      : false;
    options.isAscending = options.isAscending ? options.isAscending : true;
    options.orderBy = options.orderBy ? options.orderBy : "Id";
    return options;
  };

  useEffect(() => {
    sp.setup(context);
  }, []);

  const getListItems = async (list: string, options?: IItemOptions) => {
    /**
     * BUGS:
     * when there is no option provided on orderBy but the isAscending key is set to false, it gives no items
     */
    let listItems = [];
    setIsLoading(true);
    const newOptions = assignDefaults(options);

    try {
      if (newOptions.isId)
        if (newOptions.itemId) {
          listItems = await sp.web.lists
            .getById(list)
            .items.getById(newOptions.itemId)
            .select(newOptions.select)
            .expand(newOptions.expand)
            .get();
        } else {
          listItems = await sp.web.lists
            .getById(list)
            .items.select(newOptions.select)
            .filter(newOptions.filter)
            .expand(newOptions.expand)
            .skip(newOptions.skip, newOptions.isSkipReverse)
            .top(newOptions.top)
            .orderBy(newOptions.orderBy, newOptions.isAscending)
            .get();
        }
      else {
        if (newOptions.itemId) {
          listItems = await sp.web.lists
            .getByTitle(list)
            .items.getById(newOptions.itemId)
            .select(newOptions.select)
            .expand(newOptions.expand)
            .get();
        } else {
          listItems = await sp.web.lists
            .getByTitle(list)
            .items.select(newOptions.select)
            .filter(newOptions.filter)
            .expand(newOptions.expand)
            .skip(newOptions.skip, newOptions.isSkipReverse)
            .top(newOptions.top)
            .orderBy(newOptions.orderBy, newOptions.isAscending)
            .get();
        }
      }
    } catch (e) {
      log(e);
    }
    if (!newOptions.disableAutoRefresh) setItems(listItems);
    setIsLoading(false);

    return listItems;
  };

  const addListItemSp = async (
    list: string,
    data: Object,
    reloadWhenDone = true,
    isId?: boolean
  ) => {
    let result: IItemAddResult;
    setIsLoading(true);
    try {
      if (isId) result = await sp.web.lists.getById(list).items.add(data);
      else result = await sp.web.lists.getByTitle(list).items.add(data);

      //if (reloadWhenDone) setItems([...itemsClone, result.data]);

      setIsLoading(false);
      return result.data;
    } catch (e) {
      log(e);
      setIsLoading(false);
      return undefined;
    }
  };

  return { items, isLoading, getListItems };
};

export const useForm = (
  callback: Function,
  initialValue = {},
  validate?: Function
) => {
  const [values, setValues] = useState(initialValue);
  const [errors, setErrors] = useState({});

  const onChange = event => {
    setValues({
      ...values,
      [event.target.name]: event.target.value
    });
  };

  const customOnChange = (key: string, value: any) => {
    setValues({
      ...values,
      [key]: value
    });
  };

  const onSubmit = event => {
    event.preventDefault();

    if (!validate) {
      callback();
    } else {
      if (Object.keys(validate(values)).length === 0) {
        callback();
      } else {
        setErrors(validate(values));
      }
    }
  };

  const resetForm = () => {
    setValues(initialValue);
    setErrors({});
  };

  return {
    onChange,
    customOnChange,
    onSubmit,
    errors,
    values,
    resetForm
  };
};

export const useGlobalState = () => {
  type Action = {
    type: number;
    payload: any;
    key: string;
  };
  const dispatch: (action: Action) => void = useContext(DispatchContext);
  const state: any = useContext(StateContext);

  return { dispatch, state };
};
