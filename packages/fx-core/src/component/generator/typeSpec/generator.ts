// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Context, FxError, GeneratorResult, Inputs, ok, Result } from "@microsoft/teamsfx-api";
import { DefaultTemplateGenerator } from "../defaultGenerator";
import { TemplateNames } from "../templates/templateNames";
import { ProgrammingLanguage, QuestionNames } from "../../../question";
import { ActionContext } from "../../middleware/actionExecutionMW";
import { TemplateInfo } from "../templates/templateInfo";
import { Generator } from "../generator";

/**
 * @author bowsong@microsoft.com
 */

export class TypeSpecGenerator extends DefaultTemplateGenerator {
  componentName = "typespec-generator";
  public override activate(context: Context, inputs: Inputs): boolean {
    return [TemplateNames.DeclarativeAgentWithTypeSpec].includes(
      inputs[QuestionNames.TemplateName]
    );
  }

  public override async getTemplateInfos(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<TemplateInfo[], FxError>> {
    const appName = inputs[QuestionNames.AppName];
    const language = inputs[QuestionNames.ProgrammingLanguage] as ProgrammingLanguage;
    const safeProjectNameFromVS =
      language === "csharp" ? inputs[QuestionNames.SafeProjectName] : undefined;

    const replaceMap = {
      ...Generator.getDefaultVariables(
        appName,
        safeProjectNameFromVS,
        inputs.targetFramework,
        inputs.placeProjectFileInSolutionDir === "true"
      ),
      DeclarativeCopilot: "true",
    };
    const templateName = inputs[QuestionNames.TemplateName];

    return Promise.resolve(
      ok([
        {
          templateName,
          language: ProgrammingLanguage.None,
          replaceMap,
        },
      ])
    );
  }

  // eslint-disable-next-line @typescript-eslint/require-await
  public override async post(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<GeneratorResult, FxError>> {
    return ok({});
  }
}
