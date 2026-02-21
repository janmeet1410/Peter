import * as React from "react";
import { useEffect } from "react";

const FacebookPageEmbed = () => {
  useEffect(() => {
    // Load Facebook SDK
    const script = document.createElement("script");
    script.async = true;
    script.defer = true;
    script.crossOrigin = "anonymous";
    script.src = "https://connect.facebook.net/en_US/sdk.js#xfbml=1&version=v22.0";
    document.body.appendChild(script);
  }, []);

  return (
    <div>
      <div id="fb-root"></div>
      <div className="fb-page" data-href="https://www.facebook.com/SpecialtyRx/" data-height="500" data-small-header="false" data-adapt-container-width="true" data-hide-cover="false" data-show-facepile="false" data-show-posts="true" data-width="600">
        <blockquote cite="https://www.facebook.com/SpecialtyRx/" className="fb-xfbml-parse-ignore">
          <a href="https://www.facebook.com/SpecialtyRx/">SpecialtyRx </a>
        </blockquote>
      </div>
    </div>
  );
};

export default FacebookPageEmbed;
